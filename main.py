from kivy.config import Config

# UI config before importing most Kivy widgets
Config.set('graphics', 'width', '360')
Config.set('graphics', 'height', '640')
Config.set(
    'kivy',
    'default_font',
    [
        'Microsoft YaHei',
        r'C:\Windows\Fonts\msyh.ttc',
        r'C:\Windows\Fonts\msyh.ttc',
        r'C:\Windows\Fonts\msyh.ttc',
        r'C:\Windows\Fonts\msyh.ttc'
    ]
)

from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.button import Button
from kivy.uix.spinner import Spinner
from kivy.uix.scrollview import ScrollView
from kivy.uix.gridlayout import GridLayout
from kivy.uix.popup import Popup
from kivy.graphics import Color, Rectangle
from kivy.core.window import Window
from kivy.clock import Clock
from kivy.utils import get_color_from_hex, platform
from kivy.metrics import dp
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.uix.filechooser import FileChooserListView

import csv
import json
import os
from datetime import datetime
from openpyxl import Workbook, load_workbook

Window.size = (360, 640)


def safe_float(value, default=0.0):
    try:
        return float(value)
    except Exception:
        return default


class MainScreen(Screen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.name = 'main'

        self.records = []
        self.categories = ['饮食正餐', '娱乐消费', '学习提升', '交通', '水电', '人情世故', '房租', '医疗', '其他']
        self.records_file = os.path.join(self.get_data_dir(), 'records.json')
        self.categories_file = os.path.join(self.get_data_dir(), 'categories.json')

        if platform == 'android':
            self.request_android_permissions()

        layout = BoxLayout(orientation='vertical', padding=dp(10), spacing=dp(10))

        title = Label(
            text='个人记账',
            font_size='24sp',
            bold=True,
            size_hint=(1, 0.1),
            color=get_color_from_hex('#2c3e50')
        )
        layout.add_widget(title)

        form_layout = BoxLayout(orientation='vertical', spacing=dp(10))

        form_layout.add_widget(Label(text='消费备注:', size_hint=(1, None), height=dp(30)))
        self.name_input = TextInput(
            multiline=False,
            size_hint=(1, None),
            height=dp(40),
            background_color=get_color_from_hex('#ecf0f1'),
            foreground_color=(0, 0, 0, 1),
            cursor_color=(0, 0, 0, 1)
        )
        form_layout.add_widget(self.name_input)

        form_layout.add_widget(Label(text='分类:', size_hint=(1, None), height=dp(30)))
        self.category_spinner = Spinner(
            text='饮食正餐',
            values=self.categories,
            size_hint=(1, None),
            height=dp(40)
        )
        form_layout.add_widget(self.category_spinner)

        form_layout.add_widget(Label(text='金额(元):', size_hint=(1, None), height=dp(30)))
        self.amount_input = TextInput(
            multiline=False,
            input_filter='float',
            size_hint=(1, None),
            height=dp(40),
            background_color=get_color_from_hex('#ecf0f1'),
            foreground_color=(0, 0, 0, 1),
            cursor_color=(0, 0, 0, 1)
        )
        form_layout.add_widget(self.amount_input)

        form_layout.add_widget(Label(text='日期:', size_hint=(1, None), height=dp(30)))
        date_layout = BoxLayout(orientation='horizontal', spacing=dp(10), size_hint=(1, None), height=dp(60))
        now = datetime.now()

        year_layout = BoxLayout(orientation='vertical', spacing=dp(5))
        year_layout.add_widget(Label(text='年', size_hint=(1, None), height=dp(20)))
        self.year_input = TextInput(
            text=str(now.year),
            multiline=False,
            input_filter='int',
            size_hint=(1, None),
            height=dp(30),
            background_color=get_color_from_hex('#ecf0f1'),
            foreground_color=(0, 0, 0, 1),
            cursor_color=(0, 0, 0, 1)
        )
        year_layout.add_widget(self.year_input)
        date_layout.add_widget(year_layout)

        month_layout = BoxLayout(orientation='vertical', spacing=dp(5))
        month_layout.add_widget(Label(text='月', size_hint=(1, None), height=dp(20)))
        self.month_spinner = Spinner(
            text=str(now.month),
            values=[str(m) for m in range(1, 13)],
            size_hint=(1, None),
            height=dp(30)
        )
        month_layout.add_widget(self.month_spinner)
        date_layout.add_widget(month_layout)

        day_layout = BoxLayout(orientation='vertical', spacing=dp(5))
        day_layout.add_widget(Label(text='日', size_hint=(1, None), height=dp(20)))
        self.day_spinner = Spinner(
            text=str(now.day),
            values=[str(d) for d in range(1, 32)],
            size_hint=(1, None),
            height=dp(30)
        )
        day_layout.add_widget(self.day_spinner)
        date_layout.add_widget(day_layout)

        form_layout.add_widget(date_layout)
        layout.add_widget(form_layout)

        record_btn = Button(
            text='记录账单',
            size_hint=(1, None),
            height=dp(50),
            background_color=get_color_from_hex('#27ae60'),
            background_normal=''
        )
        record_btn.bind(on_press=self.record_bill)
        layout.add_widget(record_btn)

        button_grid = GridLayout(cols=2, spacing=dp(10), size_hint=(1, 0.28))
        buttons = [
            ('本月统计', '#9b59b6', self.show_monthly_stats),
            ('历史统计', '#3498db', self.show_history_stats),
            ('分类设置', '#e67e22', self.show_categories),
            ('导出数据', '#1abc9c', self.export_data),
            ('查看记录', '#f39c12', self.show_records),
            ('删除记录', '#e74c3c', self.delete_records),
            ('导入数据', '#2ecc71', self.import_data),
        ]
        for text, color, callback in buttons:
            btn = Button(text=text, background_color=get_color_from_hex(color), background_normal='')
            btn.bind(on_press=callback)
            button_grid.add_widget(btn)
        layout.add_widget(button_grid)

        expense_layout = BoxLayout(orientation='vertical', size_hint=(1, 0.15), padding=dp(10))
        with expense_layout.canvas.before:
            Color(rgba=get_color_from_hex('#ecf0f1')[:3] + [1])
            self.expense_rect = Rectangle(pos=expense_layout.pos, size=expense_layout.size)
        expense_layout.bind(pos=self.update_rect, size=self.update_rect)

        expense_layout.add_widget(Label(
            text='本月总支出:',
            font_size='16sp',
            bold=True,
            color=get_color_from_hex('#2c3e50')
        ))
        self.monthly_expense_label = Label(text='0.00 元', font_size='24sp', bold=True, color=get_color_from_hex('#e74c3c'))
        expense_layout.add_widget(self.monthly_expense_label)
        layout.add_widget(expense_layout)
        self.add_widget(layout)

        self.load_data()
        Clock.schedule_once(lambda dt: self.update_monthly_expense(), 0.2)

    def request_android_permissions(self):
        try:
            from android.permissions import request_permissions, Permission
            request_permissions([
                Permission.READ_EXTERNAL_STORAGE,
                Permission.WRITE_EXTERNAL_STORAGE,
            ])
        except Exception:
            pass

    def get_data_dir(self):
        data_dir = App.get_running_app().user_data_dir if App.get_running_app() else os.getcwd()
        os.makedirs(data_dir, exist_ok=True)
        return data_dir

    def get_export_dir(self):
        if platform == 'android':
            try:
                from android.storage import primary_external_storage_path
                download_dir = os.path.join(primary_external_storage_path(), 'Download')
                os.makedirs(download_dir, exist_ok=True)
                return download_dir
            except Exception:
                return self.get_data_dir()
        return os.getcwd()

    def get_import_start_dir(self):
        if platform == 'android':
            try:
                from android.storage import primary_external_storage_path
                return os.path.join(primary_external_storage_path(), 'Download')
            except Exception:
                return self.get_data_dir()
        return os.getcwd()

    def update_rect(self, instance, value):
        self.expense_rect.pos = instance.pos
        self.expense_rect.size = instance.size

    def sort_records(self):
        def sort_key(record):
            date_str = str(record.get('日期', '')).strip()
            time_str = str(record.get('记录时间', '')).strip()
            try:
                if time_str:
                    dt = datetime.strptime(time_str, '%Y-%m-%d %H:%M:%S')
                elif date_str:
                    dt = datetime.strptime(date_str, '%Y-%m-%d')
                else:
                    dt = datetime.min
            except Exception:
                try:
                    dt = datetime.strptime(date_str, '%Y-%m-%d')
                except Exception:
                    dt = datetime.min
            return dt
        self.records.sort(key=sort_key, reverse=True)

    def load_data(self):
        try:
            if os.path.exists(self.records_file):
                with open(self.records_file, 'r', encoding='utf-8') as f:
                    self.records = json.load(f)
            if os.path.exists(self.categories_file):
                with open(self.categories_file, 'r', encoding='utf-8') as f:
                    loaded_categories = json.load(f)
                    for cat in loaded_categories:
                        if cat not in self.categories:
                            self.categories.append(cat)
            self.sort_records()
            self.category_spinner.values = self.categories
            if self.categories:
                self.category_spinner.text = self.categories[0]
        except Exception:
            self.records = []
            self.category_spinner.values = self.categories

    def save_data(self):
        try:
            self.sort_records()
            with open(self.records_file, 'w', encoding='utf-8') as f:
                json.dump(self.records, f, ensure_ascii=False, indent=2)
            with open(self.categories_file, 'w', encoding='utf-8') as f:
                json.dump(self.categories, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.show_popup('错误', f'保存数据失败：\n{str(e)}')

    def record_bill(self, instance):
        name = self.name_input.text.strip()
        category = self.category_spinner.text.strip()
        amount_str = self.amount_input.text.strip()

        if not name:
            self.show_popup('错误', '请输入消费备注。')
            return
        if not amount_str:
            self.show_popup('错误', '请输入金额。')
            return

        try:
            amount = float(amount_str)
            if amount <= 0:
                raise ValueError
        except Exception:
            self.show_popup('错误', '请输入有效的正数金额。')
            return

        try:
            year_text = self.year_input.text.strip()
            if not year_text:
                self.show_popup('错误', '请输入年份。')
                return
            year = int(year_text)
            month = int(self.month_spinner.text)
            day = int(self.day_spinner.text)
            if year < 1900 or year > 9999:
                self.show_popup('错误', '请输入合理的年份，例如 2026。')
                return
            date_obj = datetime(year, month, day)
            date_str = date_obj.strftime('%Y-%m-%d')
        except Exception:
            self.show_popup('错误', '日期无效，请检查年月日。')
            return

        record = {
            '姓名/备注': name,
            '分类': category,
            '金额': amount,
            '日期': date_str,
            '记录时间': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }
        self.records.append(record)
        self.sort_records()
        self.save_data()
        self.name_input.text = ''
        self.amount_input.text = ''
        self.update_monthly_expense()
        self.show_popup('成功', f'已记录：\n{name}\n{category} - {amount:.2f}元')

    def update_monthly_expense(self):
        current_month = datetime.now().strftime('%Y-%m')
        total = 0.0
        for record in self.records:
            try:
                record_date = datetime.strptime(str(record.get('日期', '')), '%Y-%m-%d')
                amount = float(record.get('金额', 0))
                if record_date.strftime('%Y-%m') == current_month:
                    total += amount
            except Exception:
                continue
        self.monthly_expense_label.text = f'{total:.2f} 元'

    def show_monthly_stats(self, instance):
        self.show_stats_for_month(datetime.now().strftime('%Y-%m'))

    def show_history_stats(self, instance):
        months = set()
        for record in self.records:
            try:
                date_obj = datetime.strptime(str(record.get('日期', '')), '%Y-%m-%d')
                months.add(date_obj.strftime('%Y-%m'))
            except Exception:
                continue
        if not months:
            self.show_popup('提示', '暂无历史记录。')
            return
        month_list = sorted(list(months), reverse=True)
        content = BoxLayout(orientation='vertical', spacing=dp(10))
        month_spinner = Spinner(text=month_list[0], values=month_list, size_hint_y=None, height=dp(50))
        content.add_widget(month_spinner)
        open_btn = Button(text='查看统计', size_hint_y=None, height=dp(45))
        close_btn = Button(text='关闭', size_hint_y=None, height=dp(40))
        popup = Popup(title='历史统计', content=content, size_hint=(0.8, 0.5))

        def open_stats(btn):
            popup.dismiss()
            self.show_stats_for_month(month_spinner.text)

        open_btn.bind(on_press=open_stats)
        close_btn.bind(on_press=popup.dismiss)
        content.add_widget(open_btn)
        content.add_widget(close_btn)
        popup.open()

    def show_stats_for_month(self, month_str):
        month_records = []
        for record in self.records:
            try:
                date_obj = datetime.strptime(str(record.get('日期', '')), '%Y-%m-%d')
                if date_obj.strftime('%Y-%m') == month_str:
                    month_records.append(record)
            except Exception:
                continue
        if not month_records:
            self.show_popup('提示', f'{month_str} 没有记录。')
            return
        total = 0.0
        category_stats = {}
        for record in month_records:
            try:
                amount = float(record.get('金额', 0))
                category = str(record.get('分类', '未分类'))
                total += amount
                category_stats[category] = category_stats.get(category, 0) + amount
            except Exception:
                continue
        lines = [f'{month_str} 统计结果', '', f'总支出：{total:.2f} 元', '', '分类明细：']
        for category, amount in sorted(category_stats.items(), key=lambda x: x[1], reverse=True):
            lines.append(f'{category}：{amount:.2f} 元')
        self.show_popup('统计结果', '\n'.join(lines))

    def show_categories(self, instance):
        content = BoxLayout(orientation='vertical', spacing=dp(10), padding=dp(10))
        scroll = ScrollView(size_hint=(1, 1))
        grid = GridLayout(cols=1, spacing=dp(8), size_hint_y=None)
        grid.bind(minimum_height=grid.setter('height'))
        for category in self.categories:
            row = BoxLayout(size_hint_y=None, height=dp(40), spacing=dp(8))
            row.add_widget(Label(text=category))
            delete_btn = Button(text='删除', size_hint=(0.25, 1), background_color=get_color_from_hex('#e74c3c'), background_normal='')
            delete_btn.bind(on_press=lambda btn, cat=category: self.delete_category(cat))
            row.add_widget(delete_btn)
            grid.add_widget(row)
        scroll.add_widget(grid)
        content.add_widget(scroll)
        self.new_category_input = TextInput(hint_text='输入新分类', multiline=False, size_hint_y=None, height=dp(40))
        content.add_widget(self.new_category_input)
        add_btn = Button(text='添加分类', size_hint_y=None, height=dp(45), background_color=get_color_from_hex('#27ae60'), background_normal='')
        add_btn.bind(on_press=self.add_category)
        content.add_widget(add_btn)
        close_btn = Button(text='关闭', size_hint_y=None, height=dp(40))
        content.add_widget(close_btn)
        popup = Popup(title='分类设置', content=content, size_hint=(0.85, 0.85))
        close_btn.bind(on_press=popup.dismiss)
        popup.open()

    def add_category(self, instance):
        new_category = self.new_category_input.text.strip()
        if not new_category:
            self.show_popup('提示', '请输入分类名称。')
            return
        if new_category in self.categories:
            self.show_popup('提示', '该分类已存在。')
            return
        self.categories.append(new_category)
        self.category_spinner.values = self.categories
        self.save_data()
        self.show_popup('成功', f'已添加分类：{new_category}')
        self.new_category_input.text = ''

    def delete_category(self, category):
        if category not in self.categories:
            return
        if len(self.categories) <= 1:
            self.show_popup('提示', '至少保留一个分类。')
            return
        self.categories.remove(category)
        self.category_spinner.values = self.categories
        if self.category_spinner.text == category:
            self.category_spinner.text = self.categories[0]
        self.save_data()
        self.show_popup('成功', f'已删除分类：{category}')

    def show_records(self, instance):
        if not self.records:
            self.show_popup('提示', '暂无记录。')
            return
        self.sort_records()
        content = BoxLayout(orientation='vertical', spacing=dp(10), padding=dp(10))
        scroll = ScrollView(size_hint=(1, 1))
        grid = GridLayout(cols=1, spacing=dp(10), size_hint_y=None)
        grid.bind(minimum_height=grid.setter('height'))
        for record in self.records[:50]:
            note = str(record.get('姓名/备注', ''))
            category = str(record.get('分类', ''))
            amount = float(record.get('金额', 0))
            date_str = str(record.get('日期', ''))
            text = f'{date_str}  {category}\n{amount:.2f}元  {note}'
            row = Label(text=text, size_hint_y=None, height=dp(60), halign='left', valign='middle', text_size=(dp(280), None))
            grid.add_widget(row)
        scroll.add_widget(grid)
        content.add_widget(scroll)
        close_btn = Button(text='关闭', size_hint_y=None, height=dp(40))
        content.add_widget(close_btn)
        popup = Popup(title='查看记录（最近50条）', content=content, size_hint=(0.9, 0.9))
        close_btn.bind(on_press=popup.dismiss)
        popup.open()

    def delete_records(self, instance):
        if not self.records:
            self.show_popup('提示', '暂无记录可删除。')
            return
        self.sort_records()
        content = BoxLayout(orientation='vertical', spacing=dp(10), padding=dp(10))
        scroll = ScrollView(size_hint=(1, 1))
        grid = GridLayout(cols=1, spacing=dp(10), size_hint_y=None)
        grid.bind(minimum_height=grid.setter('height'))
        for real_index, record in list(enumerate(self.records[:20])):
            item_layout = BoxLayout(size_hint_y=None, height=dp(65), spacing=dp(8))
            record_text = (
                f"{record.get('日期', '')} {record.get('分类', '')}\n"
                f"{safe_float(record.get('金额', 0)):.2f}元 {str(record.get('姓名/备注', ''))[:12]}"
            )
            info_label = Label(text=record_text, size_hint=(0.72, 1), halign='left', valign='middle', text_size=(dp(190), None))
            item_layout.add_widget(info_label)
            delete_btn = Button(text='删除', size_hint=(0.28, 1), background_color=get_color_from_hex('#e74c3c'), background_normal='')
            delete_btn.bind(on_press=lambda btn, idx=real_index: self.delete_single_record(idx))
            item_layout.add_widget(delete_btn)
            grid.add_widget(item_layout)
        scroll.add_widget(grid)
        content.add_widget(scroll)
        clear_all_btn = Button(text='清空所有记录', size_hint_y=None, height=dp(40), background_color=get_color_from_hex('#c0392b'), background_normal='')
        clear_all_btn.bind(on_press=self.clear_all_records)
        content.add_widget(clear_all_btn)
        close_btn = Button(text='关闭', size_hint_y=None, height=dp(40))
        content.add_widget(close_btn)
        popup = Popup(title='删除记录（最近20条）', content=content, size_hint=(0.9, 0.9))
        close_btn.bind(on_press=popup.dismiss)
        popup.open()

    def delete_single_record(self, index):
        if 0 <= index < len(self.records):
            del self.records[index]
            self.save_data()
            self.update_monthly_expense()
            self.show_popup('成功', '记录已删除。')

    def clear_all_records(self, instance):
        def confirm_clear(btn):
            self.records = []
            self.save_data()
            self.update_monthly_expense()
            self.show_popup('成功', '所有记录已清空。')
        self.show_confirm_popup('确认清空', '确定要清空所有记录吗？此操作不可撤销。', confirm_clear)

    def export_data(self, instance):
        if not self.records:
            self.show_popup('提示', '暂无记录可导出。')
            return
        content = BoxLayout(orientation='vertical', spacing=dp(10), padding=dp(10))
        options = [
            ('导出为Excel', self.export_to_excel),
            ('导出为CSV', self.export_to_csv),
            ('导出为JSON', self.export_to_json)
        ]
        for text, callback in options:
            btn = Button(text=text, size_hint_y=None, height=dp(50))
            btn.bind(on_press=callback)
            content.add_widget(btn)
        hint = Label(text=f'默认导出目录：\n{self.get_export_dir()}', size_hint_y=None, height=dp(70), halign='center', valign='middle')
        hint.bind(size=lambda inst, val: setattr(inst, 'text_size', val))
        content.add_widget(hint)
        close_btn = Button(text='关闭', size_hint_y=None, height=dp(40))
        content.add_widget(close_btn)
        popup = Popup(title='导出数据', content=content, size_hint=(0.88, 0.72))
        close_btn.bind(on_press=popup.dismiss)
        popup.open()

    def export_to_excel(self, instance):
        try:
            self.sort_records()
            wb = Workbook()
            ws = wb.active
            ws.title = '记账记录'
            headers = ['姓名/备注', '分类', '金额', '日期', '记录时间']
            ws.append(headers)
            for record in self.records:
                ws.append([
                    record.get('姓名/备注', ''),
                    record.get('分类', ''),
                    safe_float(record.get('金额', 0)),
                    record.get('日期', ''),
                    record.get('记录时间', ''),
                ])
            filename = f'记账记录_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
            path = os.path.join(self.get_export_dir(), filename)
            wb.save(path)
            self.show_popup('成功', f'数据已导出为：\n{path}')
        except Exception as e:
            self.show_popup('错误', f'导出失败：\n{str(e)}')

    def export_to_csv(self, instance):
        try:
            self.sort_records()
            filename = f'记账记录_{datetime.now().strftime("%Y%m%d_%H%M%S")}.csv'
            path = os.path.join(self.get_export_dir(), filename)
            headers = ['姓名/备注', '分类', '金额', '日期', '记录时间']
            with open(path, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.DictWriter(f, fieldnames=headers)
                writer.writeheader()
                for record in self.records:
                    writer.writerow({key: record.get(key, '') for key in headers})
            self.show_popup('成功', f'数据已导出为：\n{path}')
        except Exception as e:
            self.show_popup('错误', f'导出失败：\n{str(e)}')

    def export_to_json(self, instance):
        try:
            self.sort_records()
            filename = f'记账记录_{datetime.now().strftime("%Y%m%d_%H%M%S")}.json'
            path = os.path.join(self.get_export_dir(), filename)
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(self.records, f, ensure_ascii=False, indent=2)
            self.show_popup('成功', f'数据已导出为：\n{path}')
        except Exception as e:
            self.show_popup('错误', f'导出失败：\n{str(e)}')

    def import_data(self, instance):
        chooser = FileChooserListView(
            path=self.get_import_start_dir(),
            filters=['*.json', '*.csv', '*.xlsx'],
            dirselect=False
        )
        content = BoxLayout(orientation='vertical', spacing=dp(10), padding=dp(10))
        content.add_widget(chooser)
        btn_layout = BoxLayout(size_hint_y=None, height=dp(45), spacing=dp(10))
        ok_btn = Button(text='导入')
        cancel_btn = Button(text='取消')
        btn_layout.add_widget(ok_btn)
        btn_layout.add_widget(cancel_btn)
        content.add_widget(btn_layout)
        popup = Popup(title='选择要导入的数据文件', content=content, size_hint=(0.95, 0.9), auto_dismiss=False)

        def do_import(btn):
            selection = chooser.selection
            if not selection:
                self.show_popup('提示', '请先选择一个文件。')
                return
            popup.dismiss()
            self.import_from_path(selection[0])

        ok_btn.bind(on_press=do_import)
        cancel_btn.bind(on_press=popup.dismiss)
        popup.open()

    def import_from_path(self, file_path):
        try:
            imported_records = []
            if file_path.lower().endswith('.json'):
                with open(file_path, 'r', encoding='utf-8') as f:
                    imported_records = json.load(f)
            elif file_path.lower().endswith('.csv'):
                with open(file_path, 'r', encoding='utf-8-sig', newline='') as f:
                    reader = csv.DictReader(f)
                    imported_records = list(reader)
            elif file_path.lower().endswith('.xlsx'):
                wb = load_workbook(file_path, data_only=True)
                ws = wb.active
                rows = list(ws.iter_rows(values_only=True))
                if not rows:
                    imported_records = []
                else:
                    headers = [str(x).strip() if x is not None else '' for x in rows[0]]
                    for row in rows[1:]:
                        record = {}
                        for idx, header in enumerate(headers):
                            if header:
                                value = row[idx] if idx < len(row) else ''
                                record[header] = '' if value is None else value
                        imported_records.append(record)
            else:
                self.show_popup('错误', '不支持的文件格式。')
                return

            if not isinstance(imported_records, list):
                self.show_popup('导入失败', '文件内容格式不正确。')
                return

            existing_keys = set()
            for record in self.records:
                try:
                    note = str(record.get('姓名/备注', '')).strip()
                    category = str(record.get('分类', '')).strip()
                    amount = round(float(record.get('金额', 0)), 2)
                    date_str = str(record.get('日期', '')).strip()
                    existing_keys.add((note, category, amount, date_str))
                except Exception:
                    continue

            valid_records = []
            new_categories = set()
            duplicate_count = 0

            for record in imported_records:
                if not isinstance(record, dict):
                    continue
                note = record.get('姓名/备注', record.get('备注', ''))
                category = record.get('分类', '')
                amount = record.get('金额', '')
                date_str = record.get('日期', '')
                record_time = record.get('记录时间', datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
                if str(note).strip() == '' or str(category).strip() == '' or str(date_str).strip() == '':
                    continue
                try:
                    amount = round(float(amount), 2)
                    if amount <= 0:
                        continue
                except Exception:
                    continue
                try:
                    if isinstance(date_str, datetime):
                        date_str = date_str.strftime('%Y-%m-%d')
                    else:
                        date_str = datetime.strptime(str(date_str)[:10], '%Y-%m-%d').strftime('%Y-%m-%d')
                except Exception:
                    try:
                        date_str = datetime.fromisoformat(str(date_str)).strftime('%Y-%m-%d')
                    except Exception:
                        continue
                clean_note = str(note).strip()
                clean_category = str(category).strip()
                key = (clean_note, clean_category, amount, date_str)
                if key in existing_keys:
                    duplicate_count += 1
                    continue
                clean_record = {
                    '姓名/备注': clean_note,
                    '分类': clean_category,
                    '金额': amount,
                    '日期': date_str,
                    '记录时间': str(record_time)
                }
                valid_records.append(clean_record)
                new_categories.add(clean_category)
                existing_keys.add(key)

            if not valid_records and duplicate_count > 0:
                self.show_popup('导入完成', f'没有新增记录。\n检测到 {duplicate_count} 条重复记录，已自动跳过。')
                return
            if not valid_records:
                self.show_popup('导入失败', '文件中没有找到可导入的有效记录。')
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
            self.show_popup('导入成功', f'成功导入 {len(valid_records)} 条记录。\n自动跳过 {duplicate_count} 条重复记录。')
        except Exception as e:
            self.show_popup('导入失败', f'发生错误：\n{str(e)}')

    def show_popup(self, title, message):
        content = BoxLayout(orientation='vertical', spacing=dp(10), padding=dp(10))
        msg_label = Label(text=message, halign='center', valign='middle')
        msg_label.bind(size=lambda instance, value: setattr(instance, 'text_size', value))
        content.add_widget(msg_label)
        ok_btn = Button(text='确定', size_hint_y=None, height=dp(40))
        content.add_widget(ok_btn)
        popup = Popup(title=title, content=content, size_hint=(0.82, 0.5), auto_dismiss=False)
        ok_btn.bind(on_press=popup.dismiss)
        popup.open()

    def show_confirm_popup(self, title, message, confirm_callback):
        content = BoxLayout(orientation='vertical', spacing=dp(10), padding=dp(10))
        msg_label = Label(text=message, halign='center', valign='middle')
        msg_label.bind(size=lambda instance, value: setattr(instance, 'text_size', value))
        content.add_widget(msg_label)
        btn_layout = BoxLayout(spacing=dp(10), size_hint_y=None, height=dp(40))
        confirm_btn = Button(text='确定')
        cancel_btn = Button(text='取消')
        btn_layout.add_widget(confirm_btn)
        btn_layout.add_widget(cancel_btn)
        content.add_widget(btn_layout)
        popup = Popup(title=title, content=content, size_hint=(0.82, 0.5), auto_dismiss=False)

        def on_confirm(btn):
            popup.dismiss()
            confirm_callback(btn)

        confirm_btn.bind(on_press=on_confirm)
        cancel_btn.bind(on_press=popup.dismiss)
        popup.open()


class AccountingApp(App):
    def build(self):
        self.title = '个人记账'
        sm = ScreenManager()
        sm.add_widget(MainScreen())
        return sm


if __name__ == '__main__':
    AccountingApp().run()
