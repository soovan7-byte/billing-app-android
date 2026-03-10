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
from kivy.graphics import Color, RoundedRectangle
from kivy.clock import Clock

from openpyxl import Workbook, load_workbook

# 桌面端设置最小窗口，安卓端不受影响
if platform in ("win", "linux", "macosx"):
    Window.minimum_width = 380
    Window.minimum_height = 700


class MainScreen(Screen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.name = "main"

        self.categories = ["饮食正餐", "娱乐消费", "学习提升", "交通", "水电", "人情世故", "房租", "医疗", "其他"]
        self.records = []

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
    # UI
    # =========================
    def build_ui(self):
        root = BoxLayout(orientation="vertical")

        scroll = ScrollView(size_hint=(1, 1))
        content = BoxLayout(
            orientation="vertical",
            spacing=dp(12),
            padding=[dp(12), dp(12), dp(12), dp(20)],
            size_hint_y=None
        )
        content.bind(minimum_height=content.setter("height"))

        # 标题
        title = Label(
            text="个人记账",
            font_size=sp(26),
            size_hint_y=None,
            height=dp(48),
            color=(0.12, 0.22, 0.36, 1)
        )
        content.add_widget(title)

        # 表单卡片
        form_card = self.make_card()
        form_layout = BoxLayout(
            orientation="vertical",
            spacing=dp(10),
            padding=dp(12),
            size_hint_y=None
        )
        form_layout.bind(minimum_height=form_layout.setter("height"))

        form_layout.add_widget(self.make_field_label("消费备注："))
        self.name_input = TextInput(
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
        form_layout.add_widget(self.name_input)

        form_layout.add_widget(self.make_field_label("分类："))
        self.category_spinner = Spinner(
            text="饮食正餐",
            values=self.categories,
            size_hint_y=None,
            height=dp(46),
            font_size=sp(18)
        )
        form_layout.add_widget(self.category_spinner)

        form_layout.add_widget(self.make_field_label("金额（元）："))
        self.amount_input = TextInput(
            multiline=False,
            input_filter="float",
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
        form_layout.add_widget(self.amount_input)

        form_layout.add_widget(self.make_field_label("日期："))

        now = datetime.now()
        date_layout = GridLayout(cols=3, spacing=dp(10), size_hint_y=None, height=dp(82))

        year_box = BoxLayout(orientation="vertical", spacing=dp(5))
        year_box.add_widget(Label(text="年", font_size=sp(16), size_hint_y=None, height=dp(24)))
        self.year_input = TextInput(
            text=str(now.year),
            multiline=False,
            input_filter="int",
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
        year_box.add_widget(self.year_input)
        date_layout.add_widget(year_box)

        month_box = BoxLayout(orientation="vertical", spacing=dp(5))
        month_box.add_widget(Label(text="月", font_size=sp(16), size_hint_y=None, height=dp(24)))
        self.month_spinner = Spinner(
            text=str(now.month),
            values=[str(i) for i in range(1, 13)],
            size_hint_y=None,
            height=dp(46),
            font_size=sp(18)
        )
        month_box.add_widget(self.month_spinner)
        date_layout.add_widget(month_box)

        day_box = BoxLayout(orientation="vertical", spacing=dp(5))
        day_box.add_widget(Label(text="日", font_size=sp(16), size_hint_y=None, height=dp(24)))
        self.day_spinner = Spinner(
            text=str(now.day),
            values=[str(i) for i in range(1, 32)],
            size_hint_y=None,
            height=dp(46),
            font_size=sp(18)
        )
        day_box.add_widget(self.day_spinner)
        date_layout.add_widget(day_box)

        form_layout.add_widget(date_layout)

        record_btn = Button(
            text="记录账单",
            size_hint_y=None,
            height=dp(52),
            font_size=sp(20),
            background_normal="",
            background_color=(0.16, 0.69, 0.37, 1),
            color=(1, 1, 1, 1)
        )
        record_btn.bind(on_press=self.record_bill)
        form_layout.add_widget(record_btn)

        form_card.add_widget(form_layout)
        content.add_widget(form_card)

        # 功能按钮区
        button_card = self.make_card()
        button_grid = GridLayout(
            cols=2,
            spacing=dp(10),
            padding=dp(12),
            size_hint_y=None
        )
        button_grid.bind(minimum_height=button_grid.setter("height"))

        buttons = [
            ("本月统计", (0.62, 0.35, 0.84, 1), self.show_monthly_stats),
            ("历史统计", (0.20, 0.60, 0.95, 1), self.show_history_stats),
            ("分类设置", (0.91, 0.52, 0.17, 1), self.show_categories),
            ("导出数据", (0.10, 0.74, 0.70, 1), self.export_data),
            ("查看记录", (0.95, 0.63, 0.10, 1), self.show_records),
            ("删除记录", (0.91, 0.30, 0.24, 1), self.delete_records),
            ("导入数据", (0.18, 0.80, 0.44, 1), self.import_data_popup),
        ]

        for text, color, callback in buttons:
            btn = Button(
                text=text,
                size_hint_y=None,
                height=dp(48),
                font_size=sp(18),
                background_normal="",
                background_color=color,
                color=(1, 1, 1, 1)
            )
            btn.bind(on_press=callback)
            button_grid.add_widget(btn)

        button_card.add_widget(button_grid)
        content.add_widget(button_card)

        # 本月总支出
        expense_card = self.make_card()
        expense_layout = BoxLayout(
            orientation="vertical",
            padding=dp(12),
            spacing=dp(6),
            size_hint_y=None,
            height=dp(110)
        )

        expense_layout.add_widget(Label(
            text="本月总支出：",
            font_size=sp(18),
            size_hint_y=None,
            height=dp(28),
            color=(0.12, 0.22, 0.36, 1)
        ))

        self.monthly_expense_label = Label(
            text="0.00 元",
            font_size=sp(30),
            size_hint_y=None,
            height=dp(50),
            color=(0.90, 0.30, 0.24, 1)
        )
        expense_layout.add_widget(self.monthly_expense_label)

        expense_card.add_widget(expense_layout)
        content.add_widget(expense_card)

        scroll.add_widget(content)
        root.add_widget(scroll)
        self.add_widget(root)

    def make_card(self):
        card = BoxLayout(size_hint_y=None)
        card.bind(minimum_height=card.setter("height"))

        with card.canvas.before:
            Color(1, 1, 1, 1)
            card.bg = RoundedRectangle(radius=[dp(16)] * 4, pos=card.pos, size=card.size)

        def update_bg(instance, value):
            instance.bg.pos = instance.pos
            instance.bg.size = instance.size

        card.bind(pos=update_bg, size=update_bg)
        return card

    def make_field_label(self, text):
        label = Label(
            text=text,
            font_size=sp(18),
            size_hint_y=None,
            height=dp(32),
            halign="left",
            valign="middle",
            color=(0.12, 0.22, 0.36, 1)
        )
        label.bind(size=lambda inst, val: setattr(inst, "text_size", (val[0], None)))
        return label

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

        for record in self.records:
            try:
                record_date = datetime.strptime(str(record.get("日期", "")), "%Y-%m-%d")
                if record_date.strftime("%Y-%m") == current_month:
                    total += float(record.get("金额", 0))
            except Exception:
                continue

        self.monthly_expense_label.text = f"{total:.2f} 元"

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
        self.show_stats_for_month(current_month)

    def show_history_stats(self, instance):
        months = set()
        for record in self.records:
            try:
                date_obj = datetime.strptime(str(record.get("日期", "")), "%Y-%m-%d")
                months.add(date_obj.strftime("%Y-%m"))
            except Exception:
                continue

        if not months:
            self.show_popup("提示", "暂无历史记录。")
            return

        month_list = sorted(list(months), reverse=True)

        content = BoxLayout(orientation="vertical", spacing=dp(12), padding=dp(12))
        spinner = Spinner(
            text=month_list[0],
            values=month_list,
            size_hint_y=None,
            height=dp(48),
            font_size=sp(18)
        )
        content.add_widget(spinner)

        btn_view = Button(text="查看统计", size_hint_y=None, height=dp(46), font_size=sp(18))
        btn_close = Button(text="关闭", size_hint_y=None, height=dp(42), font_size=sp(17))

        popup = Popup(title="历史统计", content=content, size_hint=(0.88, 0.42), auto_dismiss=False)

        def do_view(btn):
            popup.dismiss()
            self.show_stats_for_month(spinner.text)

        btn_view.bind(on_press=do_view)
        btn_close.bind(on_press=popup.dismiss)

        content.add_widget(btn_view)
        content.add_widget(btn_close)
        popup.open()

    def show_stats_for_month(self, month_str):
        month_records = []
        for record in self.records:
            try:
                date_obj = datetime.strptime(str(record.get("日期", "")), "%Y-%m-%d")
                if date_obj.strftime("%Y-%m") == month_str:
                    month_records.append(record)
            except Exception:
                continue

        if not month_records:
            self.show_popup("提示", f"{month_str} 没有记录。")
            return

        total = 0.0
        category_stats = {}

        for record in month_records:
            try:
                amount = float(record.get("金额", 0))
                category = str(record.get("分类", "未分类"))
                total += amount
                category_stats[category] = category_stats.get(category, 0.0) + amount
            except Exception:
                continue

        lines = [f"{month_str} 统计结果", "", f"总支出：{total:.2f} 元", "", "分类明细："]
        for category, amount in sorted(category_stats.items(), key=lambda x: x[1], reverse=True):
            lines.append(f"{category}：{amount:.2f} 元")

        self.show_popup("统计结果", "\n".join(lines))

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
                background_color=(0.91, 0.30, 0.24, 1),
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
            background_color=(0.16, 0.69, 0.37, 1),
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
                background_color=(0.91, 0.30, 0.24, 1),
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
            background_color=(0.75, 0.22, 0.19, 1),
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

    def export_to_excel(self, popup=None):
        try:
            self.sort_records()
            export_dir = self.get_export_dir()
            filename = f"记账记录_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            file_path = os.path.join(export_dir, filename)

            wb = Workbook()
            ws = wb.active
            ws.title = "记账记录"
            ws.append(["姓名/备注", "分类", "金额", "日期", "记录时间"])

            for record in self.records:
                ws.append([
                    record.get("姓名/备注", ""),
                    record.get("分类", ""),
                    record.get("金额", ""),
                    record.get("日期", ""),
                    record.get("记录时间", "")
                ])

            wb.save(file_path)

            if popup:
                popup.dismiss()
            self.show_popup("成功", f"数据已导出为：\n{file_path}")
        except Exception as e:
            self.show_popup("错误", f"导出失败：\n{str(e)}")

    def export_to_csv(self, popup=None):
        try:
            self.sort_records()
            export_dir = self.get_export_dir()
            filename = f"记账记录_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            file_path = os.path.join(export_dir, filename)

            fieldnames = ["姓名/备注", "分类", "金额", "日期", "记录时间"]
            with open(file_path, "w", encoding="utf-8-sig", newline="") as f:
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

            if popup:
                popup.dismiss()
            self.show_popup("成功", f"数据已导出为：\n{file_path}")
        except Exception as e:
            self.show_popup("错误", f"导出失败：\n{str(e)}")

    def export_to_json(self, popup=None):
        try:
            self.sort_records()
            export_dir = self.get_export_dir()
            filename = f"记账记录_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
            file_path = os.path.join(export_dir, filename)

            with open(file_path, "w", encoding="utf-8") as f:
                json.dump(self.records, f, ensure_ascii=False, indent=2)

            if popup:
                popup.dismiss()
            self.show_popup("成功", f"数据已导出为：\n{file_path}")
        except Exception as e:
            self.show_popup("错误", f"导出失败：\n{str(e)}")

    # =========================
    # 导入
    # =========================
    def import_data_popup(self, instance):
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
