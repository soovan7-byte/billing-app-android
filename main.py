# -*- coding: utf-8 -*-
import os
import json
import csv
from datetime import datetime

from kivy.config import Config
from kivy.utils import platform

# =========================
# 字体设置
# =========================
# 电脑端：优先用 Windows 楷体
# 安卓端：用项目目录里的 simkai.ttf
APP_DIR = os.path.dirname(os.path.abspath(__file__))
WINDOWS_FONT_PATH = r"C:\Windows\Fonts\simkai.ttf"
LOCAL_FONT_PATH = os.path.join(APP_DIR, "simkai.ttf")

font_path = None
if platform == "win" and os.path.exists(WINDOWS_FONT_PATH):
    font_path = WINDOWS_FONT_PATH
elif os.path.exists(LOCAL_FONT_PATH):
    font_path = LOCAL_FONT_PATH

if font_path:
    Config.set(
        "kivy",
        "default_font",
        ["AppFont", font_path, font_path, font_path, font_path]
    )

from kivy.app import App
from kivy.core.window import Window
from kivy.metrics import dp, sp
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.filechooser import FileChooserListView
from kivy.uix.gridlayout import GridLayout
from kivy.uix.label import Label
from kivy.uix.popup import Popup
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.uix.scrollview import ScrollView
from kivy.uix.spinner import Spinner
from kivy.uix.textinput import TextInput
from kivy.uix.widget import Widget
from kivy.graphics import Color, Ellipse, RoundedRectangle
from kivy.clock import Clock

from openpyxl import Workbook, load_workbook

# 桌面端设置最小窗口，安卓端不受影响
if platform in ("win", "linux", "macosx"):
    Window.minimum_width = 380
    Window.minimum_height = 700


CATEGORY_CHART_COLORS = [
    (0.91, 0.30, 0.24, 1),
    (0.20, 0.60, 0.86, 1),
    (0.18, 0.80, 0.44, 1),
    (0.95, 0.61, 0.07, 1),
    (0.61, 0.35, 0.71, 1),
    (0.10, 0.74, 0.61, 1),
    (0.90, 0.49, 0.13, 1),
    (0.20, 0.29, 0.37, 1),
    (0.84, 0.15, 0.45, 1),
    (0.40, 0.70, 0.20, 1),
]


class CategoryPieChart(Widget):
    """使用 Kivy Canvas 绘制按分类消费占比的正圆扇形图。"""
    def __init__(self, category_stats=None, colors=None, **kwargs):
        super().__init__(**kwargs)
        self.category_stats = category_stats or []
        self.colors = colors or CATEGORY_CHART_COLORS
        self.bind(pos=self._redraw, size=self._redraw)
        self._redraw()

    def set_data(self, category_stats):
        self.category_stats = category_stats or []
        self._redraw()

    def _redraw(self, *args):
        self.canvas.clear()
        valid_stats = [(category, amount) for category, amount in self.category_stats if amount > 0]
        total = sum(amount for category, amount in valid_stats)
        if total <= 0:
            return

        diameter = min(self.width, self.height)
        if diameter <= 0:
            return

        padding = dp(6)
        diameter = max(0, diameter - padding * 2)
        x = self.x + (self.width - diameter) / 2
        y = self.y + (self.height - diameter) / 2

        start_angle = 0
        with self.canvas:
            for index, (category, amount) in enumerate(valid_stats):
                end_angle = 360 if index == len(valid_stats) - 1 else start_angle + amount / total * 360
                Color(*self.colors[index % len(self.colors)])
                Ellipse(pos=(x, y), size=(diameter, diameter), angle_start=start_angle, angle_end=end_angle)
                start_angle = end_angle


class MainScreen(Screen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.name = "main"

        self.categories = ["饮食正餐", "娱乐消费", "学习提升", "交通", "水电", "人情世故", "房租", "医疗", "其他"]
        self.records = []
        self._android_export_bound = False
        self._pending_android_export = None
        self._pending_android_import = False

        self.storage_dir = self.get_storage_dir()
        os.makedirs(self.storage_dir, exist_ok=True)

        self.records_path = os.path.join(self.storage_dir, "records.json")
        self.categories_path = os.path.join(self.storage_dir, "categories.json")

        self.build_ui()
        self.load_data()
        Clock.schedule_once(lambda dt: self.update_monthly_expense(), 0.1)

    # =========================
    # 基础路径
    # =========================
    def get_storage_dir(self):
        app = App.get_running_app()
        if platform == "android" and app is not None:
            return app.user_data_dir
        return APP_DIR

    def get_export_dir(self):
        # 安卓优先导出到 Download；失败则回退到 app 私有目录
        if platform == "android":
            try:
                from android.storage import primary_external_storage_path
                export_dir = os.path.join(primary_external_storage_path(), "Download")
                os.makedirs(export_dir, exist_ok=True)
                return export_dir
            except Exception:
                pass

        export_dir = os.path.join(self.storage_dir, "exports")
        os.makedirs(export_dir, exist_ok=True)
        return export_dir

    def get_default_import_dir(self):
        export_dir = self.get_export_dir()
        if os.path.exists(export_dir):
            return export_dir
        return self.storage_dir

    # =========================
    # UI 辅助方法
    # =========================
    def _make_card(self, bg_color=(1, 1, 1, 1), radius=dp(16)):
        """创建一个带圆角白色背景的卡片容器"""
        card = BoxLayout(size_hint_y=None)
        card.bind(minimum_height=card.setter("height"))

        with card.canvas.before:
            Color(*bg_color)
            card.bg = RoundedRectangle(radius=[radius] * 4, pos=card.pos, size=card.size)

        def update_bg(instance, value):
            instance.bg.pos = instance.pos
            instance.bg.size = instance.size

        card.bind(pos=update_bg, size=update_bg)
        return card

    def _make_title_label(self, text, font_size=sp(24), height=dp(48)):
        label = Label(
            text=text,
            font_size=font_size,
            size_hint_y=None,
            height=height,
            color=(0.12, 0.22, 0.36, 1)
        )
        return label

    def _make_section_label(self, text, font_size=sp(18), height=dp(32)):
        label = Label(
            text=text,
            font_size=font_size,
            size_hint_y=None,
            height=height,
            halign="left",
            valign="middle",
            color=(0.12, 0.22, 0.36, 1)
        )
        label.bind(size=lambda inst, val: setattr(inst, "text_size", (val[0], None)))
        return label

    def _make_primary_button(self, text, height=dp(48), font_size=sp(18)):
        btn = Button(
            text=text,
            size_hint_y=None,
            height=height,
            font_size=font_size,
            background_normal="",
            background_color=(0.18, 0.63, 0.35, 1),
            color=(1, 1, 1, 1)
        )
        return btn

    def _make_secondary_button(self, text, height=dp(44), font_size=sp(17)):
        btn = Button(
            text=text,
            size_hint_y=None,
            height=height,
            font_size=font_size,
            background_normal="",
            background_color=(0.35, 0.55, 0.75, 1),
            color=(1, 1, 1, 1)
        )
        return btn

    def _make_action_button(self, text, color, height=dp(48), font_size=sp(18)):
        btn = Button(
            text=text,
            size_hint_y=None,
            height=height,
            font_size=font_size,
            background_normal="",
            background_color=color,
            color=(1, 1, 1, 1)
        )
        return btn

    def _make_text_input(self, hint_text="", input_filter=None, height=dp(46), font_size=sp(18), **kwargs):
        ti = TextInput(
            hint_text=hint_text,
            multiline=False,
            input_filter=input_filter,
            size_hint_y=None,
            height=height,
            font_size=font_size,
            background_normal="",
            background_active="",
            background_color=(0.96, 0.96, 0.96, 1),
            foreground_color=(0, 0, 0, 1),
            cursor_color=(0, 0, 0, 1),
            padding=[dp(10), dp(10), dp(10), dp(10)],
            disabled=False,
            readonly=False,
            write_tab=False,
            **kwargs
        )
        # 移除自定义圆角背景，使用默认 TextInput 外观
        return ti

    def _make_spinner(self, text, values, height=dp(46), font_size=sp(18)):
        sp = Spinner(
            text=text,
            values=values,
            size_hint_y=None,
            height=height,
            font_size=font_size,
            background_normal="",
            background_color=(0.10, 0.17, 0.24, 1),
            color=(1, 1, 1, 1)
        )
        # 添加圆角背景
        with sp.canvas.before:
            Color(0.85, 0.90, 0.95, 1)  # 浅蓝灰
            sp.bg = RoundedRectangle(radius=[dp(8)]*4, pos=sp.pos, size=sp.size)

        def update_bg(instance, value):
            instance.bg.pos = instance.pos
            instance.bg.size = instance.size

        sp.bind(pos=update_bg, size=update_bg)
        return sp

    # =========================
    # UI
    # =========================
    def build_ui(self):
        # 根布局：浅灰背景
        root = BoxLayout(orientation="vertical")
        with root.canvas.before:
            Color(0.94, 0.94, 0.94, 1)  # 浅灰
            root.bg = RoundedRectangle(radius=[0]*4, pos=root.pos, size=root.size)

        def update_root_bg(instance, value):
            instance.bg.pos = instance.pos
            instance.bg.size = instance.size
        root.bind(pos=update_root_bg, size=update_root_bg)

        scroll = ScrollView(size_hint=(1, 1))
        content = BoxLayout(
            orientation="vertical",
            spacing=dp(16),
            padding=[dp(16), dp(16), dp(16), dp(24)],
            size_hint_y=None
        )
        content.bind(minimum_height=content.setter("height"))


        # ===== 本月统计卡片 =====
        expense_card = self._make_card()
        expense_layout = BoxLayout(
            orientation="vertical",
            padding=[dp(16), dp(24), dp(16), dp(16)],
            spacing=dp(8),
            size_hint_y=None,
            height=dp(150)
        )

        expense_layout.add_widget(self._make_section_label("本月总支出", font_size=sp(18), height=dp(28)))

        self.monthly_expense_label = Label(
            text="0.00 元",
            font_size=sp(34),
            size_hint_y=None,
            height=dp(56),
            color=(0.90, 0.30, 0.24, 1)
        )
        expense_layout.add_widget(self.monthly_expense_label)

        # 记录数量
        self.record_count_label = Label(
            text="记录数：0",
            font_size=sp(16),
            size_hint_y=None,
            height=dp(24),
            color=(0.4, 0.4, 0.4, 1)
        )
        expense_layout.add_widget(self.record_count_label)

        expense_card.add_widget(expense_layout)
        content.add_widget(expense_card)

        # ===== 表单卡片 =====
        form_card = self._make_card()
        form_layout = BoxLayout(
            orientation="vertical",
            spacing=dp(12),
            padding=[dp(16), dp(16), dp(16), dp(16)],
            size_hint_y=None
        )
        form_layout.bind(minimum_height=form_layout.setter("height"))

        form_layout.add_widget(self._make_section_label("消费备注："))
        self.name_input = self._make_text_input(hint_text="例如：午餐")
        form_layout.add_widget(self.name_input)

        form_layout.add_widget(self._make_section_label("分类："))
        self.category_spinner = self._make_spinner(
            text="饮食正餐",
            values=self.categories
        )
        form_layout.add_widget(self.category_spinner)

        form_layout.add_widget(self._make_section_label("金额（元）："))
        self.amount_input = self._make_text_input(
            hint_text="0.00",
            input_filter="float"
        )
        form_layout.add_widget(self.amount_input)

        form_layout.add_widget(self._make_section_label("日期："))

        now = datetime.now()
        date_layout = GridLayout(cols=3, spacing=dp(10), size_hint_y=None, height=dp(82))

        year_box = BoxLayout(orientation="vertical", spacing=dp(4))
        year_box.add_widget(Label(text="年", font_size=sp(16), size_hint_y=None, height=dp(24)))
        self.year_input = self._make_text_input(
            text=str(now.year),
            input_filter="int"
        )
        year_box.add_widget(self.year_input)
        date_layout.add_widget(year_box)

        month_box = BoxLayout(orientation="vertical", spacing=dp(4))
        month_box.add_widget(Label(text="月", font_size=sp(16), size_hint_y=None, height=dp(24)))
        self.month_spinner = self._make_spinner(
            text=str(now.month),
            values=[str(i) for i in range(1, 13)]
        )
        month_box.add_widget(self.month_spinner)
        date_layout.add_widget(month_box)

        day_box = BoxLayout(orientation="vertical", spacing=dp(4))
        day_box.add_widget(Label(text="日", font_size=sp(16), size_hint_y=None, height=dp(24)))
        self.day_spinner = self._make_spinner(
            text=str(now.day),
            values=[str(i) for i in range(1, 32)]
        )
        day_box.add_widget(self.day_spinner)
        date_layout.add_widget(day_box)

        form_layout.add_widget(date_layout)

        record_btn = self._make_primary_button("记录账单", height=dp(52), font_size=sp(20))
        record_btn.bind(on_press=self.record_bill)
        form_layout.add_widget(record_btn)

        form_card.add_widget(form_layout)
        content.add_widget(form_card)


        # ===== 功能按钮卡片 =====
        button_card = self._make_card()
        button_grid = GridLayout(
            cols=2,
            spacing=dp(12),
            padding=[dp(16), dp(16), dp(16), dp(16)],
            size_hint_y=None
        )
        button_grid.bind(minimum_height=button_grid.setter("height"))

        buttons = [
            ("本月统计", (0.35, 0.56, 0.75, 1), self.show_monthly_stats),
            ("历史统计", (0.35, 0.56, 0.75, 1), self.show_history_stats),
            ("分类设置", (0.35, 0.56, 0.75, 1), self.show_categories),
            ("导出数据", (0.35, 0.56, 0.75, 1), self.export_data),
            ("查看记录", (0.35, 0.56, 0.75, 1), self.show_records),
            ("删除记录", (0.80, 0.30, 0.25, 1), self.delete_records),
            ("导入数据", (0.35, 0.56, 0.75, 1), self.import_data_popup),
        ]

        for text, color, callback in buttons:
            btn = self._make_action_button(text, color, height=dp(48), font_size=sp(17))
            btn.bind(on_press=callback)
            button_grid.add_widget(btn)

        button_card.add_widget(button_grid)
        content.add_widget(button_card)

        scroll.add_widget(content)
        root.add_widget(scroll)
        self.add_widget(root)

    def make_card(self):
        """保留原有方法，供其他可能调用的地方使用"""
        return self._make_card()

    def make_field_label(self, text):
        """保留原有方法，供其他可能调用的地方使用"""
        return self._make_section_label(text)

    # =========================
    # 数据处理
    # =========================
    def sort_records(self):
        def sort_key(record):
            record_time = str(record.get("记录时间", "")).strip()
            date_str = str(record.get("日期", "")).strip()

            try:
                if record_time:
                    return datetime.strptime(record_time, "%Y-%m-%d %H:%M:%S")
            except Exception:
                pass

            try:
                if date_str:
                    return datetime.strptime(date_str, "%Y-%m-%d")
            except Exception:
                pass

            return datetime.min

        self.records.sort(key=sort_key, reverse=True)

    def load_data(self):
        try:
            if os.path.exists(self.records_path):
                with open(self.records_path, "r", encoding="utf-8") as f:
                    self.records = json.load(f)

            if os.path.exists(self.categories_path):
                with open(self.categories_path, "r", encoding="utf-8") as f:
                    loaded_categories = json.load(f)
                    for cat in loaded_categories:
                        if cat not in self.categories:
                            self.categories.append(cat)

            self.sort_records()
            self.category_spinner.values = self.categories
            if self.categories:
                self.category_spinner.text = self.categories[0]
        except Exception as e:
            self.records = []
            self.show_popup("提示", f"读取本地数据失败：\n{str(e)}")

    def save_data(self):
        try:
            self.sort_records()
            with open(self.records_path, "w", encoding="utf-8") as f:
                json.dump(self.records, f, ensure_ascii=False, indent=2)

            with open(self.categories_path, "w", encoding="utf-8") as f:
                json.dump(self.categories, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.show_popup("错误", f"保存数据失败：\n{str(e)}")

    def update_monthly_expense(self):
        current_month = datetime.now().strftime("%Y-%m")
        total = 0.0
        count = 0

        for record in self.records:
            try:
                record_date = datetime.strptime(str(record.get("日期", "")), "%Y-%m-%d")
                if record_date.strftime("%Y-%m") == current_month:
                    total += float(record.get("金额", 0))
                    count += 1
            except Exception:
                continue

        self.monthly_expense_label.text = f"{total:.2f} 元"
        self.record_count_label.text = f"记录数：{count}"

    # =========================
    # 记账
    # =========================
    def record_bill(self, instance):
        note = self.name_input.text.strip()
        category = self.category_spinner.text.strip()
        amount_text = self.amount_input.text.strip()

        if not note:
            self.show_popup("错误", "请输入消费备注。")
            return

        if not amount_text:
            self.show_popup("错误", "请输入金额。")
            return

        try:
            amount = round(float(amount_text), 2)
            if amount <= 0:
                raise ValueError
        except Exception:
            self.show_popup("错误", "请输入有效的正数金额。")
            return

        try:
            year_text = self.year_input.text.strip()
            if not year_text:
                self.show_popup("错误", "请输入年份。")
                return

            year = int(year_text)
            month = int(self.month_spinner.text)
            day = int(self.day_spinner.text)

            if year < 1900 or year > 9999:
                self.show_popup("错误", "请输入合理的年份，例如 2026。")
                return

            date_obj = datetime(year, month, day)
            date_str = date_obj.strftime("%Y-%m-%d")
        except Exception:
            self.show_popup("错误", "日期无效，请检查年月日。")
            return

        record = {
            "姓名/备注": note,
            "分类": category,
            "金额": amount,
            "日期": date_str,
            "记录时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }

        self.records.append(record)
        self.save_data()
        self.update_monthly_expense()

        self.name_input.text = ""
        self.amount_input.text = ""

        self.show_popup("成功", f"已记录：\n{note}\n{category} - {amount:.2f}元")

    # =========================
    # 统计
    # =========================
    def show_monthly_stats(self, instance):
        current_month = datetime.now().strftime("%Y-%m")
        self.show_monthly_stats_popup(current_month)

    def show_history_stats(self, instance):
        months = self.get_available_months()
        years = self.get_available_years()

        if not months and not years:
            self.show_popup("提示", "暂无历史记录。")
            return

        content = BoxLayout(orientation="vertical", spacing=dp(12), padding=dp(12))

        type_label = Label(
            text="统计类型",
            font_size=sp(16),
            size_hint_y=None,
            height=dp(26),
            halign="left",
            valign="middle",
            color=(0.12, 0.22, 0.36, 1)
        )
        type_label.bind(size=lambda inst, val: setattr(inst, "text_size", (val[0], val[1])))
        content.add_widget(type_label)

        type_spinner = Spinner(
            text="按月统计" if months else "按年统计",
            values=("按月统计", "按年统计"),
            size_hint_y=None,
            height=dp(48),
            font_size=sp(18)
        )
        content.add_widget(type_spinner)

        period_label = Label(
            text="统计周期",
            font_size=sp(16),
            size_hint_y=None,
            height=dp(26),
            halign="left",
            valign="middle",
            color=(0.12, 0.22, 0.36, 1)
        )
        period_label.bind(size=lambda inst, val: setattr(inst, "text_size", (val[0], val[1])))
        content.add_widget(period_label)

        initial_periods = months if type_spinner.text == "按月统计" else years
        period_spinner = Spinner(
            text=initial_periods[0] if initial_periods else "",
            values=initial_periods,
            size_hint_y=None,
            height=dp(48),
            font_size=sp(18)
        )
        content.add_widget(period_spinner)

        btn_view = Button(text="查看统计", size_hint_y=None, height=dp(46), font_size=sp(18))
        btn_close = Button(text="关闭", size_hint_y=None, height=dp(42), font_size=sp(17))

        popup = Popup(title="历史统计", content=content, size_hint=(0.90, 0.52), auto_dismiss=False)

        def update_periods(spinner, text):
            periods = months if text == "按月统计" else years
            period_spinner.values = periods
            period_spinner.text = periods[0] if periods else ""

        def do_view(btn):
            if not period_spinner.text:
                self.show_popup("提示", "暂无历史记录。")
                return
            popup.dismiss()
            if type_spinner.text == "按年统计":
                self.show_stats_for_period("year", period_spinner.text)
            else:
                self.show_stats_for_period("month", period_spinner.text)

        type_spinner.bind(text=update_periods)
        btn_view.bind(on_press=do_view)
        btn_close.bind(on_press=popup.dismiss)

        content.add_widget(btn_view)
        content.add_widget(btn_close)
        popup.open()

    def _get_record_date(self, record):
        try:
            return datetime.strptime(str(record.get("日期", "")), "%Y-%m-%d")
        except Exception:
            return None

    def _get_record_amount(self, record):
        try:
            amount = float(record.get("金额", 0))
            return amount if amount > 0 else None
        except Exception:
            return None

    def get_available_months(self):
        months = set()
        for record in self.records:
            date_obj = self._get_record_date(record)
            if date_obj is not None:
                months.add(date_obj.strftime("%Y-%m"))
        return sorted(months, reverse=True)

    def get_available_years(self):
        years = set()
        for record in self.records:
            date_obj = self._get_record_date(record)
            if date_obj is not None:
                years.add(date_obj.strftime("%Y"))
        return sorted(years, reverse=True)

    def get_records_for_month(self, month_str):
        month_records = []
        for record in self.records:
            date_obj = self._get_record_date(record)
            if date_obj is not None and date_obj.strftime("%Y-%m") == month_str:
                month_records.append(record)
        return month_records

    def get_records_for_year(self, year_str):
        year_records = []
        for record in self.records:
            date_obj = self._get_record_date(record)
            if date_obj is not None and date_obj.strftime("%Y") == year_str:
                year_records.append(record)
        return year_records

    def get_category_stats(self, records):
        category_stats = {}
        for record in records:
            amount = self._get_record_amount(record)
            if amount is None:
                continue
            category = str(record.get("分类", "未分类")).strip() or "未分类"
            category_stats[category] = category_stats.get(category, 0.0) + amount
        return sorted(category_stats.items(), key=lambda x: x[1], reverse=True)

    def get_monthly_totals_for_year(self, year_str):
        monthly_totals = {month: 0.0 for month in range(1, 13)}
        for record in self.get_records_for_year(year_str):
            date_obj = self._get_record_date(record)
            amount = self._get_record_amount(record)
            if date_obj is None or amount is None:
                continue
            monthly_totals[date_obj.month] += amount
        return monthly_totals

    def _make_stats_label(self, text, font_size, height, color, halign="center"):
        label = Label(
            text=text,
            font_size=font_size,
            size_hint_y=None,
            height=height,
            halign=halign,
            valign="middle",
            color=color
        )
        label.bind(size=lambda inst, val: setattr(inst, "text_size", (val[0], val[1])))
        return label

    def _make_category_detail_row(self, category, amount, total, color):
        row = BoxLayout(orientation="horizontal", size_hint_y=None, height=dp(54), spacing=dp(8), padding=(0, dp(4)))

        marker_box = BoxLayout(size_hint=(None, 1), width=dp(18), padding=(0, dp(16), 0, dp(16)))
        marker = Widget(size_hint=(None, None), size=(dp(14), dp(14)))
        with marker.canvas.before:
            Color(*color)
            marker.color_box = RoundedRectangle(radius=[dp(3)] * 4, pos=marker.pos, size=marker.size)

        def update_marker(instance, value):
            instance.color_box.pos = instance.pos
            instance.color_box.size = instance.size

        marker.bind(pos=update_marker, size=update_marker)
        marker_box.add_widget(marker)
        row.add_widget(marker_box)

        category_label = Label(
            text=category,
            font_size=sp(15),
            size_hint=(1, None),
            height=dp(46),
            halign="left",
            valign="middle",
            color=(0.12, 0.12, 0.12, 1)
        )
        category_label.bind(size=lambda inst, val: setattr(inst, "text_size", (val[0], val[1])))
        row.add_widget(category_label)

        percent = amount / total * 100 if total > 0 else 0
        amount_label = Label(
            text=f"{amount:.2f} 元\n{percent:.1f}%",
            font_size=sp(14),
            size_hint=(None, None),
            width=dp(104),
            height=dp(46),
            halign="right",
            valign="middle",
            color=(0.12, 0.12, 0.12, 1)
        )
        amount_label.bind(size=lambda inst, val: setattr(inst, "text_size", (val[0], val[1])))
        row.add_widget(amount_label)
        return row

    def _build_category_stats_section(self, category_stats, total, empty_text):
        section = BoxLayout(orientation="vertical", spacing=dp(6), size_hint_y=None)
        section.bind(minimum_height=section.setter("height"))

        if total > 0:
            chart = CategoryPieChart(category_stats, size_hint_y=None, height=dp(190))
            section.add_widget(chart)
        else:
            section.add_widget(self._make_stats_label(
                empty_text,
                sp(17),
                dp(110),
                (0.45, 0.45, 0.45, 1)
            ))

        section.add_widget(self._make_stats_label(
            "分类明细",
            sp(17),
            dp(28),
            (0.12, 0.22, 0.36, 1),
            halign="left"
        ))

        if total > 0:
            for index, (category, amount) in enumerate(category_stats):
                section.add_widget(self._make_category_detail_row(
                    category, amount, total, CATEGORY_CHART_COLORS[index % len(CATEGORY_CHART_COLORS)]
                ))
        else:
            section.add_widget(self._make_stats_label(
                empty_text,
                sp(16),
                dp(52),
                (0.45, 0.45, 0.45, 1)
            ))
        return section

    def _make_month_total_row(self, month, amount):
        row = BoxLayout(orientation="horizontal", size_hint_y=None, height=dp(38), spacing=dp(8))
        month_label = Label(
            text=f"{month}月",
            font_size=sp(15),
            size_hint=(0.35, None),
            height=dp(38),
            halign="left",
            valign="middle",
            color=(0.12, 0.12, 0.12, 1)
        )
        month_label.bind(size=lambda inst, val: setattr(inst, "text_size", (val[0], val[1])))
        row.add_widget(month_label)

        amount_label = Label(
            text=f"{amount:.2f}元",
            font_size=sp(15),
            size_hint=(0.65, None),
            height=dp(38),
            halign="right",
            valign="middle",
            color=(0.12, 0.12, 0.12, 1)
        )
        amount_label.bind(size=lambda inst, val: setattr(inst, "text_size", (val[0], val[1])))
        row.add_widget(amount_label)
        return row

    def _build_year_monthly_totals_section(self, monthly_totals):
        section = BoxLayout(orientation="vertical", spacing=dp(4), size_hint_y=None)
        section.bind(minimum_height=section.setter("height"))
        section.add_widget(self._make_stats_label(
            "1月至12月支出明细",
            sp(17),
            dp(30),
            (0.12, 0.22, 0.36, 1),
            halign="left"
        ))
        for month in range(1, 13):
            section.add_widget(self._make_month_total_row(month, monthly_totals.get(month, 0.0)))
        return section

    def show_monthly_stats_popup(self, month_str):
        self.show_stats_for_period("current_month", month_str)

    def show_stats_for_month(self, month_str):
        self.show_stats_for_period("month", month_str)

    def show_stats_for_period(self, period_type, period_value):
        is_year = period_type == "year"
        is_current_month = period_type == "current_month"

        if is_year:
            records = self.get_records_for_year(period_value)
            category_stats = self.get_category_stats(records)
            total = sum(amount for category, amount in category_stats)
            monthly_totals = self.get_monthly_totals_for_year(period_value)
            title_text = f"{period_value} 年统计"
            total_text = f"年度总支出：{total:.2f} 元" if total > 0 else f"{period_value} 年没有有效消费记录"
            empty_text = f"{period_value} 年没有有效消费记录"
            popup_title = "年度统计"
        else:
            records = self.get_records_for_month(period_value)
            category_stats = self.get_category_stats(records)
            total = sum(amount for category, amount in category_stats)
            monthly_totals = None
            title_text = f"{period_value} 本月统计" if is_current_month else f"{period_value} 月统计"
            total_prefix = "本月总支出" if is_current_month else "月度总支出"
            total_text = f"{total_prefix}：{total:.2f} 元" if total > 0 else f"{period_value} 没有有效消费记录"
            empty_text = "本月暂无消费记录" if is_current_month else f"{period_value} 没有有效消费记录"
            popup_title = "本月统计" if is_current_month else "月度统计"

        content = BoxLayout(orientation="vertical", spacing=dp(10), padding=dp(12))
        content.add_widget(self._make_stats_label(
            title_text,
            sp(20),
            dp(34),
            (0.12, 0.22, 0.36, 1)
        ))
        content.add_widget(self._make_stats_label(
            total_text,
            sp(18),
            dp(32),
            (0.18, 0.45, 0.25, 1)
        ))

        scroll = ScrollView(size_hint=(1, 1), do_scroll_x=False)
        scroll_content = BoxLayout(orientation="vertical", spacing=dp(10), size_hint_y=None)
        scroll_content.bind(minimum_height=scroll_content.setter("height"))
        scroll_content.add_widget(self._build_category_stats_section(category_stats, total, empty_text))

        if is_year:
            scroll_content.add_widget(self._build_year_monthly_totals_section(monthly_totals))

        scroll.add_widget(scroll_content)
        content.add_widget(scroll)

        btn_close = Button(text="关闭", size_hint_y=None, height=dp(44), font_size=sp(17))
        content.add_widget(btn_close)

        popup = Popup(title=popup_title, content=content, size_hint=(0.94, 0.90), auto_dismiss=False)
        btn_close.bind(on_press=popup.dismiss)
        popup.open()

    # =========================
    # 分类设置
    # =========================
    def show_categories(self, instance):
        content = BoxLayout(orientation="vertical", spacing=dp(10), padding=dp(10))

        scroll = ScrollView(size_hint=(1, 1))
        grid = GridLayout(cols=1, spacing=dp(8), size_hint_y=None)
        grid.bind(minimum_height=grid.setter("height"))

        for category in self.categories:
            row = BoxLayout(size_hint_y=None, height=dp(48), spacing=dp(8))
            row.add_widget(Label(text=category, font_size=sp(17), halign="left", valign="middle"))

            delete_btn = Button(
                text="删除",
                size_hint=(0.28, 1),
                font_size=sp(16),
                background_normal="",
                background_color=(0.80, 0.30, 0.25, 1),
                color=(1, 1, 1, 1)
            )
            delete_btn.bind(on_press=lambda btn, cat=category: self.delete_category(cat))
            row.add_widget(delete_btn)

            grid.add_widget(row)

        scroll.add_widget(grid)
        content.add_widget(scroll)

        self.new_category_input = TextInput(
            hint_text="输入新分类",
            multiline=False,
            size_hint_y=None,
            height=dp(46),
            font_size=sp(18),
            background_normal="",
            background_active="",
            background_color=(0.96, 0.96, 0.96, 1),
            foreground_color=(0, 0, 0, 1),
            cursor_color=(0, 0, 0, 1),
            padding=[dp(10), dp(10), dp(10), dp(10)]
        )
        content.add_widget(self.new_category_input)

        add_btn = Button(
            text="添加分类",
            size_hint_y=None,
            height=dp(46),
            font_size=sp(18),
            background_normal="",
            background_color=(0.20, 0.60, 0.35, 1),
            color=(1, 1, 1, 1)
        )
        add_btn.bind(on_press=self.add_category)
        content.add_widget(add_btn)

        close_btn = Button(text="关闭", size_hint_y=None, height=dp(42), font_size=sp(17))
        content.add_widget(close_btn)

        popup = Popup(title="分类设置", content=content, size_hint=(0.9, 0.9), auto_dismiss=False)
        close_btn.bind(on_press=popup.dismiss)
        popup.open()

    def add_category(self, instance):
        new_category = self.new_category_input.text.strip()
        if not new_category:
            self.show_popup("提示", "请输入分类名称。")
            return

        if new_category in self.categories:
            self.show_popup("提示", "该分类已存在。")
            return

        self.categories.append(new_category)
        self.category_spinner.values = self.categories
        self.save_data()
        self.new_category_input.text = ""
        self.show_popup("成功", f"已添加分类：{new_category}")

    def delete_category(self, category):
        if category not in self.categories:
            return

        if len(self.categories) <= 1:
            self.show_popup("提示", "至少保留一个分类。")
            return

        self.categories.remove(category)
        self.category_spinner.values = self.categories
        if self.category_spinner.text == category and self.categories:
            self.category_spinner.text = self.categories[0]

        self.save_data()
        self.show_popup("成功", f"已删除分类：{category}")

    # =========================
    # 查看记录
    # =========================
    def show_records(self, instance):
        if not self.records:
            self.show_popup("提示", "暂无记录。")
            return

        self.sort_records()

        content = BoxLayout(orientation="vertical", spacing=dp(10), padding=dp(10))
        scroll = ScrollView(size_hint=(1, 1))
        grid = GridLayout(cols=1, spacing=dp(8), size_hint_y=None)
        grid.bind(minimum_height=grid.setter("height"))

        display_records = self.records[:50]

        for record in display_records:
            note = str(record.get("姓名/备注", ""))
            category = str(record.get("分类", ""))
            amount = float(record.get("金额", 0))
            date_str = str(record.get("日期", ""))

            text = f"{date_str}  {category}\n{amount:.2f}元  {note}"
            row = Label(
                text=text,
                font_size=sp(16),
                size_hint_y=None,
                height=dp(62),
                halign="left",
                valign="middle"
            )
            row.bind(size=lambda inst, val: setattr(inst, "text_size", (val[0] - dp(10), None)))
            grid.add_widget(row)

        scroll.add_widget(grid)
        content.add_widget(scroll)

        close_btn = Button(text="关闭", size_hint_y=None, height=dp(42), font_size=sp(17))
        content.add_widget(close_btn)

        popup = Popup(
            title="查看记录（最近50条）",
            content=content,
            size_hint=(0.92, 0.9),
            auto_dismiss=False
        )
        close_btn.bind(on_press=popup.dismiss)
        popup.open()

    # =========================
    # 删除记录
    # =========================
    def delete_records(self, instance):
        if not self.records:
            self.show_popup("提示", "暂无记录可删除。")
            return

        self.sort_records()

        content = BoxLayout(orientation="vertical", spacing=dp(10), padding=dp(10))
        scroll = ScrollView(size_hint=(1, 1))
        grid = GridLayout(cols=1, spacing=dp(8), size_hint_y=None)
        grid.bind(minimum_height=grid.setter("height"))

        display_records = list(enumerate(self.records[:20]))

        for real_index, record in display_records:
            row = BoxLayout(size_hint_y=None, height=dp(68), spacing=dp(8))

            record_text = (
                f"{record.get('日期', '')} {record.get('分类', '')}\n"
                f"{float(record.get('金额', 0)):.2f}元 {str(record.get('姓名/备注', ''))[:14]}"
            )

            info_label = Label(
                text=record_text,
                font_size=sp(15),
                size_hint=(0.72, 1),
                halign="left",
                valign="middle"
            )
            info_label.bind(size=lambda inst, val: setattr(inst, "text_size", (val[0] - dp(6), None)))
            row.add_widget(info_label)

            delete_btn = Button(
                text="删除",
                font_size=sp(16),
                size_hint=(0.28, 1),
                background_normal="",
                background_color=(0.80, 0.30, 0.25, 1),
                color=(1, 1, 1, 1)
            )
            delete_btn.bind(on_press=lambda btn, idx=real_index: self.delete_single_record(idx))
            row.add_widget(delete_btn)

            grid.add_widget(row)

        scroll.add_widget(grid)
        content.add_widget(scroll)

        clear_btn = Button(
            text="清空所有记录",
            size_hint_y=None,
            height=dp(44),
            font_size=sp(17),
            background_normal="",
            background_color=(0.80, 0.30, 0.25, 1),
            color=(1, 1, 1, 1)
        )
        clear_btn.bind(on_press=self.clear_all_records)
        content.add_widget(clear_btn)

        close_btn = Button(text="关闭", size_hint_y=None, height=dp(42), font_size=sp(17))
        content.add_widget(close_btn)

        popup = Popup(
            title="删除记录（最近20条）",
            content=content,
            size_hint=(0.92, 0.9),
            auto_dismiss=False
        )
        close_btn.bind(on_press=popup.dismiss)
        popup.open()

    def delete_single_record(self, index):
        if 0 <= index < len(self.records):
            del self.records[index]
            self.save_data()
            self.update_monthly_expense()
            self.show_popup("成功", "记录已删除。")

    def clear_all_records(self, instance):
        def do_clear(btn):
            self.records = []
            self.save_data()
            self.update_monthly_expense()
            self.show_popup("成功", "所有记录已清空。")

        self.show_confirm_popup("确认清空", "确定要清空所有记录吗？此操作不可撤销。", do_clear)

    # =========================
    # 导出
    # =========================
    def export_data(self, instance):
        if not self.records:
            self.show_popup("提示", "暂无记录可导出。")
            return

        content = BoxLayout(orientation="vertical", spacing=dp(10), padding=dp(12))

        btn_xlsx = Button(text="导出为 Excel", size_hint_y=None, height=dp(48), font_size=sp(18))
        btn_csv = Button(text="导出为 CSV", size_hint_y=None, height=dp(48), font_size=sp(18))
        btn_json = Button(text="导出为 JSON", size_hint_y=None, height=dp(48), font_size=sp(18))
        btn_close = Button(text="关闭", size_hint_y=None, height=dp(42), font_size=sp(17))

        popup = Popup(title="导出数据", content=content, size_hint=(0.86, 0.48), auto_dismiss=False)

        btn_xlsx.bind(on_press=lambda btn: self.export_to_excel(popup))
        btn_csv.bind(on_press=lambda btn: self.export_to_csv(popup))
        btn_json.bind(on_press=lambda btn: self.export_to_json(popup))
        btn_close.bind(on_press=popup.dismiss)

        content.add_widget(btn_xlsx)
        content.add_widget(btn_csv)
        content.add_widget(btn_json)
        content.add_widget(btn_close)

        popup.open()

    def _is_android(self):
        return platform == "android"

    def _get_export_fieldnames(self):
        return ["姓名/备注", "分类", "金额", "日期", "记录时间"]

    def _get_export_timestamp(self):
        return datetime.now().strftime("%Y%m%d_%H%M%S")

    def _format_exception(self, error):
        return f"{type(error).__name__}: {str(error)}"

    def _create_excel_temp_file(self, timestamp):
        temp_path = os.path.join(self.storage_dir, f"export_temp_{timestamp}.xlsx")
        wb = Workbook()
        ws = wb.active
        ws.title = "记账记录"
        ws.append(self._get_export_fieldnames())

        for record in self.records:
            ws.append([
                record.get("姓名/备注", ""),
                record.get("分类", ""),
                record.get("金额", ""),
                record.get("日期", ""),
                record.get("记录时间", "")
            ])

        wb.save(temp_path)
        return temp_path

    def _create_csv_temp_file(self, timestamp):
        temp_path = os.path.join(self.storage_dir, f"export_temp_{timestamp}.csv")
        fieldnames = self._get_export_fieldnames()
        with open(temp_path, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            for record in self.records:
                writer.writerow({
                    "姓名/备注": record.get("姓名/备注", ""),
                    "分类": record.get("分类", ""),
                    "金额": record.get("金额", ""),
                    "日期": record.get("日期", ""),
                    "记录时间": record.get("记录时间", "")
                })
        return temp_path

    def _create_json_temp_file(self, timestamp):
        temp_path = os.path.join(self.storage_dir, f"export_temp_{timestamp}.json")
        with open(temp_path, "w", encoding="utf-8") as f:
            json.dump(self.records, f, ensure_ascii=False, indent=2)
        return temp_path

    def _create_export_temp_file(self, export_type, timestamp):
        if export_type == "excel":
            return self._create_excel_temp_file(timestamp)
        if export_type == "csv":
            return self._create_csv_temp_file(timestamp)
        if export_type == "json":
            return self._create_json_temp_file(timestamp)
        raise ValueError(f"未知导出类型：{export_type}")

    def _cleanup_export_temp_file(self, temp_path):
        if not temp_path:
            return None
        try:
            if os.path.exists(temp_path):
                os.remove(temp_path)
        except Exception as e:
            return self._format_exception(e)
        return None

    def _get_android_export_config(self, export_type, timestamp):
        configs = {
            "excel": {
                "request_code": 5101,
                "extension": "xlsx",
                "mime_type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "success_message": "Excel 导出成功"
            },
            "csv": {
                "request_code": 5102,
                "extension": "csv",
                "mime_type": "text/csv",
                "success_message": "CSV 导出成功"
            },
            "json": {
                "request_code": 5103,
                "extension": "json",
                "mime_type": "application/json",
                "success_message": "JSON 导出成功"
            }
        }
        config = configs[export_type].copy()
        config["filename"] = f"记账记录_{timestamp}.{config['extension']}"
        return config

    def _ensure_android_export_binding(self):
        if not self._is_android() or self._android_export_bound:
            return
        if platform == "android":
            from android import activity
            activity.bind(on_activity_result=self._on_android_activity_result)
            self._android_export_bound = True

    def _start_android_document_export(self, export_type, popup=None):
        if not self.records:
            self.show_popup("提示", "暂无记录可导出。")
            return

        if self._pending_android_export is not None:
            self.show_popup("提示", "已有导出任务正在等待系统保存界面返回，请稍后再试。")
            return

        temp_path = None
        try:
            self.sort_records()
            timestamp = self._get_export_timestamp()
            config = self._get_android_export_config(export_type, timestamp)
            temp_path = self._create_export_temp_file(export_type, timestamp)
            self._ensure_android_export_binding()

            if platform == "android":
                from jnius import autoclass
                Intent = autoclass("android.content.Intent")
                PythonActivity = autoclass("org.kivy.android.PythonActivity")

                intent = Intent(Intent.ACTION_CREATE_DOCUMENT)
                intent.addCategory(Intent.CATEGORY_OPENABLE)
                intent.setType(config["mime_type"])
                intent.putExtra(Intent.EXTRA_TITLE, config["filename"])

                self._pending_android_export = {
                    "type": export_type,
                    "request_code": config["request_code"],
                    "temp_path": temp_path,
                    "filename": config["filename"],
                    "success_message": config["success_message"]
                }

                if popup:
                    popup.dismiss()
                PythonActivity.mActivity.startActivityForResult(intent, config["request_code"])
        except Exception as e:
            cleanup_warning = self._cleanup_export_temp_file(temp_path)
            self._pending_android_export = None
            message = f"导出失败：\n{self._format_exception(e)}"
            if cleanup_warning:
                message += f"\n临时文件清理警告：{cleanup_warning}"
            self.show_popup("错误", message)

    def _on_android_activity_result(self, request_code, result_code, intent):
        Clock.schedule_once(
            lambda dt: self._dispatch_android_activity_result(request_code, result_code, intent),
            0
        )

    def _dispatch_android_activity_result(self, request_code, result_code, intent):
        if request_code == 5201:
            self._handle_android_import_result(result_code, intent)
            return
        self._handle_android_activity_result(request_code, result_code, intent)

    def _handle_android_activity_result(self, request_code, result_code, intent):
        pending = self._pending_android_export
        if not pending or request_code != pending.get("request_code"):
            return

        temp_path = pending.get("temp_path")
        try:
            if platform == "android":
                from jnius import autoclass
                Activity = autoclass("android.app.Activity")

                if result_code != Activity.RESULT_OK:
                    return

                if intent is None:
                    self.show_popup("错误", "导出失败：\nAndroid 系统保存界面未返回数据。")
                    return

                uri = intent.getData()
                if uri is None:
                    self.show_popup("错误", "导出失败：\nAndroid 系统保存界面未返回文件 URI。")
                    return

                self._write_file_to_content_uri(temp_path, uri)
                self.show_popup("成功", pending.get("success_message", "导出成功"))
        except Exception as e:
            self.show_popup("错误", f"导出失败：\n{self._format_exception(e)}")
        finally:
            cleanup_warning = self._cleanup_export_temp_file(temp_path)
            self._pending_android_export = None
            if cleanup_warning:
                print(f"导出临时文件清理失败：{cleanup_warning}")

    def _write_file_to_content_uri(self, source_path, uri):
        if platform != "android":
            raise RuntimeError("Content URI 写入仅支持 Android 平台。")
        if not source_path or not os.path.exists(source_path):
            raise FileNotFoundError(f"临时导出文件不存在：{source_path}")

        from jnius import autoclass
        PythonActivity = autoclass("org.kivy.android.PythonActivity")
        resolver = PythonActivity.mActivity.getContentResolver()
        output_stream = resolver.openOutputStream(uri)
        if output_stream is None:
            raise OSError("ContentResolver.openOutputStream 返回 None。")

        try:
            with open(source_path, "rb") as source_file:
                while True:
                    chunk = source_file.read(64 * 1024)
                    if not chunk:
                        break
                    output_stream.write(chunk)
            output_stream.flush()
        finally:
            output_stream.close()

    def export_to_excel(self, popup=None):
        if self._is_android():
            self._start_android_document_export("excel", popup)
            return
        try:
            self.sort_records()
            export_dir = self.get_export_dir()
            timestamp = self._get_export_timestamp()
            filename = f"记账记录_{timestamp}.xlsx"
            file_path = os.path.join(export_dir, filename)
            temp_path = self._create_excel_temp_file(timestamp)
            try:
                os.replace(temp_path, file_path)
            except Exception:
                self._cleanup_export_temp_file(temp_path)
                raise

            if popup:
                popup.dismiss()
            self.show_popup("成功", f"数据已导出为：\n{file_path}")
        except Exception as e:
            self.show_popup("错误", f"导出失败：\n{self._format_exception(e)}")

    def export_to_csv(self, popup=None):
        if self._is_android():
            self._start_android_document_export("csv", popup)
            return
        try:
            self.sort_records()
            export_dir = self.get_export_dir()
            timestamp = self._get_export_timestamp()
            filename = f"记账记录_{timestamp}.csv"
            file_path = os.path.join(export_dir, filename)
            temp_path = self._create_csv_temp_file(timestamp)
            try:
                os.replace(temp_path, file_path)
            except Exception:
                self._cleanup_export_temp_file(temp_path)
                raise

            if popup:
                popup.dismiss()
            self.show_popup("成功", f"数据已导出为：\n{file_path}")
        except Exception as e:
            self.show_popup("错误", f"导出失败：\n{self._format_exception(e)}")

    def export_to_json(self, popup=None):
        if self._is_android():
            self._start_android_document_export("json", popup)
            return
        try:
            self.sort_records()
            export_dir = self.get_export_dir()
            timestamp = self._get_export_timestamp()
            filename = f"记账记录_{timestamp}.json"
            file_path = os.path.join(export_dir, filename)
            temp_path = self._create_json_temp_file(timestamp)
            try:
                os.replace(temp_path, file_path)
            except Exception:
                self._cleanup_export_temp_file(temp_path)
                raise

            if popup:
                popup.dismiss()
            self.show_popup("成功", f"数据已导出为：\n{file_path}")
        except Exception as e:
            self.show_popup("错误", f"导出失败：\n{self._format_exception(e)}")

    # =========================
    # 导入
    # =========================
    def import_data_popup(self, instance):
        if self._is_android():
            self._start_android_document_import()
            return

        chooser = FileChooserListView(
            path=self.get_default_import_dir(),
            filters=["*.json", "*.csv", "*.xlsx"],
            size_hint=(1, 1)
        )

        content = BoxLayout(orientation="vertical", spacing=dp(10), padding=dp(10))
        content.add_widget(chooser)

        btn_box = BoxLayout(size_hint_y=None, height=dp(44), spacing=dp(10))
        btn_import = Button(text="导入", font_size=sp(17))
        btn_cancel = Button(text="取消", font_size=sp(17))
        btn_box.add_widget(btn_import)
        btn_box.add_widget(btn_cancel)
        content.add_widget(btn_box)

        popup = Popup(title="选择要导入的数据文件", content=content, size_hint=(0.94, 0.92), auto_dismiss=False)

        def do_import(btn):
            if not chooser.selection:
                self.show_popup("提示", "请先选择一个文件。")
                return
            file_path = chooser.selection[0]
            popup.dismiss()
            self.import_file(file_path)

        btn_import.bind(on_press=do_import)
        btn_cancel.bind(on_press=popup.dismiss)

        popup.open()

    def _start_android_document_import(self):
        if self._pending_android_import:
            self.show_popup("提示", "文件选择器已打开，请先完成或取消当前选择。")
            return

        try:
            self._ensure_android_export_binding()
            if platform == "android":
                from jnius import autoclass
                Intent = autoclass("android.content.Intent")
                PythonActivity = autoclass("org.kivy.android.PythonActivity")

                intent = Intent(Intent.ACTION_OPEN_DOCUMENT)
                intent.addCategory(Intent.CATEGORY_OPENABLE)
                intent.setType("*/*")
                intent.putExtra(Intent.EXTRA_ALLOW_MULTIPLE, False)
                intent.putExtra(Intent.EXTRA_MIME_TYPES, [
                    "application/json",
                    "text/json",
                    "text/csv",
                    "application/csv",
                    "text/comma-separated-values",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "application/octet-stream",
                ])
                self._pending_android_import = True
                PythonActivity.mActivity.startActivityForResult(intent, 5201)
        except Exception as e:
            self._pending_android_import = False
            self.show_popup("错误", f"无法打开系统文件选择器：\n{self._format_exception(e)}")

    def _handle_android_import_result(self, result_code, intent):
        if not self._pending_android_import:
            return

        temp_path = None
        try:
            if platform == "android":
                from jnius import autoclass
                Activity = autoclass("android.app.Activity")

                if result_code != Activity.RESULT_OK:
                    return
                if intent is None:
                    raise ValueError("Android 系统文件选择器未返回数据。")

                uri = intent.getData()
                if uri is None:
                    raise ValueError("Android 系统文件选择器未返回文件 URI。")

                filename, mime_type = self._get_android_document_info(uri)
                extension = self._resolve_import_extension(filename, mime_type)
                if extension is None:
                    self.show_popup("错误", "无法识别文件格式，请选择 JSON、CSV 或 XLSX 文件。")
                    return

                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
                temp_path = os.path.join(self.storage_dir, f"import_temp_{timestamp}.{extension}")
                self._copy_content_uri_to_file(uri, temp_path)
                self.import_file(temp_path)
        except Exception as e:
            self.show_popup("导入失败", f"读取所选文件失败：\n{self._format_exception(e)}")
        finally:
            self._pending_android_import = False
            cleanup_warning = self._cleanup_export_temp_file(temp_path)
            if cleanup_warning:
                print(f"导入临时文件清理失败：{cleanup_warning}")

    def _get_android_document_info(self, uri):
        from jnius import autoclass
        PythonActivity = autoclass("org.kivy.android.PythonActivity")
        OpenableColumns = autoclass("android.provider.OpenableColumns")
        resolver = PythonActivity.mActivity.getContentResolver()
        mime_type = resolver.getType(uri)
        filename = None
        cursor = None
        try:
            cursor = resolver.query(uri, None, None, None, None)
            if cursor is not None and cursor.moveToFirst():
                name_index = cursor.getColumnIndex(OpenableColumns.DISPLAY_NAME)
                if name_index >= 0 and not cursor.isNull(name_index):
                    filename = str(cursor.getString(name_index))
        except Exception as e:
            print(f"读取所选文件名称失败：{self._format_exception(e)}")
        finally:
            if cursor is not None:
                cursor.close()
        return filename, str(mime_type) if mime_type else None

    def _resolve_import_extension(self, filename, mime_type):
        supported_extensions = {".json": "json", ".csv": "csv", ".xlsx": "xlsx"}
        if filename:
            extension = os.path.splitext(filename)[1].lower()
            if extension in supported_extensions:
                return supported_extensions[extension]

        mime_extensions = {
            "application/json": "json",
            "text/json": "json",
            "text/csv": "csv",
            "application/csv": "csv",
            "text/comma-separated-values": "csv",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": "xlsx",
        }
        return mime_extensions.get((mime_type or "").lower())

    def _copy_content_uri_to_file(self, uri, target_path):
        if platform != "android":
            raise RuntimeError("Content URI 读取仅支持 Android 平台。")

        from jnius import autoclass
        PythonActivity = autoclass("org.kivy.android.PythonActivity")
        resolver = PythonActivity.mActivity.getContentResolver()
        input_stream = resolver.openInputStream(uri)
        if input_stream is None:
            raise OSError("ContentResolver.openInputStream 返回 None。")

        try:
            buffer = bytearray(64 * 1024)
            with open(target_path, "wb") as target_file:
                while True:
                    count = input_stream.read(buffer)
                    if count == -1:
                        break
                    if count > 0:
                        target_file.write(buffer[:count])
        finally:
            input_stream.close()

    def import_file(self, file_path):
        try:
            imported_records = []

            if file_path.lower().endswith(".json"):
                with open(file_path, "r", encoding="utf-8") as f:
                    imported_records = json.load(f)

            elif file_path.lower().endswith(".csv"):
                with open(file_path, "r", encoding="utf-8-sig", newline="") as f:
                    reader = csv.DictReader(f)
                    imported_records = list(reader)

            elif file_path.lower().endswith(".xlsx"):
                wb = load_workbook(file_path, data_only=True)
                ws = wb.active
                rows = list(ws.iter_rows(values_only=True))
                if not rows:
                    self.show_popup("导入失败", "Excel 文件为空。")
                    return

                headers = [str(h).strip() if h is not None else "" for h in rows[0]]
                for row in rows[1:]:
                    item = {}
                    for i, header in enumerate(headers):
                        if header:
                            item[header] = row[i] if i < len(row) else ""
                    imported_records.append(item)

            else:
                self.show_popup("错误", "不支持的文件格式。")
                return

            if not isinstance(imported_records, list):
                self.show_popup("导入失败", "文件内容格式不正确。")
                return

            existing_keys = set()
            for record in self.records:
                try:
                    note = str(record.get("姓名/备注", "")).strip()
                    category = str(record.get("分类", "")).strip()
                    amount = round(float(record.get("金额", 0)), 2)
                    date_str = str(record.get("日期", "")).strip()
                    key = (note, category, amount, date_str)
                    existing_keys.add(key)
                except Exception:
                    continue

            valid_records = []
            new_categories = set()
            duplicate_count = 0

            for record in imported_records:
                if not isinstance(record, dict):
                    continue

                note = record.get("姓名/备注", record.get("备注", ""))
                category = record.get("分类", "")
                amount = record.get("金额", "")
                date_str = record.get("日期", "")
                record_time = record.get("记录时间", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

                if str(note).strip() == "" or str(category).strip() == "" or str(date_str).strip() == "":
                    continue

                try:
                    amount = round(float(amount), 2)
                    if amount <= 0:
                        continue
                except Exception:
                    continue

                try:
                    if isinstance(date_str, datetime):
                        date_str = date_str.strftime("%Y-%m-%d")
                    else:
                        date_str = datetime.strptime(str(date_str)[:10], "%Y-%m-%d").strftime("%Y-%m-%d")
                except Exception:
                    continue

                clean_note = str(note).strip()
                clean_category = str(category).strip()
                key = (clean_note, clean_category, amount, date_str)

                if key in existing_keys:
                    duplicate_count += 1
                    continue

                clean_record = {
                    "姓名/备注": clean_note,
                    "分类": clean_category,
                    "金额": amount,
                    "日期": date_str,
                    "记录时间": str(record_time)
                }

                valid_records.append(clean_record)
                new_categories.add(clean_category)
                existing_keys.add(key)

            if not valid_records and duplicate_count > 0:
                self.show_popup("导入完成", f"没有新增记录。\n检测到 {duplicate_count} 条重复记录，已自动跳过。")
                return

            if not valid_records:
                self.show_popup("导入失败", "文件中没有找到可导入的有效记录。")
                return

            self.records.extend(valid_records)
            self.sort_records()

            for cat in sorted(new_categories):
                if cat and cat not in self.categories:
                    self.categories.append(cat)

            self.category_spinner.values = self.categories
            if self.category_spinner.text not in self.categories and self.categories:
                self.category_spinner.text = self.categories[0]

            self.save_data()
            self.update_monthly_expense()

            self.show_popup(
                "导入成功",
                f"成功导入 {len(valid_records)} 条记录。\n自动跳过 {duplicate_count} 条重复记录。"
            )

        except Exception as e:
            self.show_popup("导入失败", f"发生错误：\n{str(e)}")

    # =========================
    # 通用弹窗
    # =========================
    def show_popup(self, title, message):
        content = BoxLayout(orientation="vertical", spacing=dp(10), padding=dp(12))

        msg = Label(
            text=message,
            font_size=sp(17),
            halign="center",
            valign="middle"
        )
        msg.bind(size=lambda inst, val: setattr(inst, "text_size", (val[0] - dp(8), None)))
        content.add_widget(msg)

        btn = Button(text="确定", size_hint_y=None, height=dp(42), font_size=sp(17))
        content.add_widget(btn)

        popup = Popup(title=title, content=content, size_hint=(0.86, 0.42), auto_dismiss=False)
        btn.bind(on_press=popup.dismiss)
        popup.open()

    def show_confirm_popup(self, title, message, confirm_callback):
        content = BoxLayout(orientation="vertical", spacing=dp(10), padding=dp(12))

        msg = Label(
            text=message,
            font_size=sp(17),
            halign="center",
            valign="middle"
        )
        msg.bind(size=lambda inst, val: setattr(inst, "text_size", (val[0] - dp(8), None)))
        content.add_widget(msg)

        btn_box = BoxLayout(size_hint_y=None, height=dp(42), spacing=dp(10))
        btn_ok = Button(text="确定", font_size=sp(17))
        btn_cancel = Button(text="取消", font_size=sp(17))
        btn_box.add_widget(btn_ok)
        btn_box.add_widget(btn_cancel)
        content.add_widget(btn_box)

        popup = Popup(title=title, content=content, size_hint=(0.86, 0.42), auto_dismiss=False)

        def do_confirm(btn):
            popup.dismiss()
            confirm_callback(btn)

        btn_ok.bind(on_press=do_confirm)
        btn_cancel.bind(on_press=popup.dismiss)
        popup.open()


class AccountingApp(App):
    def build(self):
        self.title = "个人记账"
        sm = ScreenManager()
        sm.add_widget(MainScreen())
        return sm


if __name__ == "__main__":
    AccountingApp().run()
