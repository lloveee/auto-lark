from core.env import EXCEL_DIR
from datetime import datetime
from modes.mode_cha_lvyue import get_table_filter
from modes.mode_cha_lvyue import get_table_value
from core.logger import logger
from modes.excel.excel_tool import ExcelTool
from pathlib import Path
import flet as ft
import os


class ExcelPreviewControl(ft.Column):
    """Excel 预览控件"""

    def __init__(self, excel_tool: ExcelTool, page: ft.Page):
        super().__init__()
        self.excel_tool = excel_tool
        self._page = page  # 在初始化时接收 page 引用
        self.sheet_dropdown = None
        self.data_table = None
        self.file_picker = None
        self.select_file_button = None
        self.remote_load_btn = None
        self.refresh_button = None
        self.scrollable_table = None
        self.open_button = None

        # 立即构建 UI
        self._build_ui()

    def _build_ui(self):
        """构建 UI 组件"""
        self.sheet_dropdown = ft.Dropdown(
            label="选择工作表",
            options=[ft.dropdown.Option(sheet) for sheet in self.excel_tool.get_sheet_names()],
            value=self.excel_tool.get_sheet_names()[0] if self.excel_tool.get_sheet_names() else None
        )

        self.data_table = ft.DataTable(
            columns=[],
            rows=[],
            border=ft.border.all(1, ft.Colors.GREY_400),
            heading_row_color=ft.Colors.GREY_200,
            expand=True,
            column_spacing=10,
        )

        # 使用传入的 page 引用
        self.file_picker = ft.FilePicker(on_result=self.on_file_picked)
        self._page.overlay.append(self.file_picker)

        self.remote_load_btn = ft.ElevatedButton(
            text="加载远程数据",
            icon=ft.Icons.SEARCH,
            style=ft.ButtonStyle(
                bgcolor=ft.Colors.BLUE_400,
                color=ft.Colors.WHITE,
                shape=ft.RoundedRectangleBorder(radius=10),
                padding=ft.padding.symmetric(horizontal=20, vertical=10),
            ),
            on_click=self.on_load_data_from_remote,
        )

        self.refresh_button = ft.ElevatedButton(
            text="刷新",
            on_click=lambda e: self.update_table(self.sheet_dropdown.value),
            icon=ft.Icons.REFRESH
        )

        self.select_file_button = ft.ElevatedButton(
            text="选择 需要查看的 Excel 文件",
            on_click=lambda e: self.file_picker.pick_files(
                allowed_extensions=["xlsx", "xls"],
                allow_multiple=False
            ),
            icon=ft.Icons.FOLDER_OPEN
        )

        self.open_button = ft.ElevatedButton(
            text="打开 Excel 文件",
            on_click=self.open_excel_file,
            icon=ft.Icons.FILE_OPEN
        )

        self.scrollable_table = ft.Container(
            content=ft.Column(
                controls=[
                    ft.Row(
                        scroll=ft.ScrollMode.ALWAYS,
                        controls=[self.data_table],
                    )
                ],
                scroll=ft.ScrollMode.ALWAYS,
                expand=True,
            )
        )

        # 设置 Column 的 controls
        self.controls = [
            ft.Row(
                [self.sheet_dropdown, self.select_file_button, self.refresh_button,
                 self.open_button, self.remote_load_btn],
                alignment=ft.MainAxisAlignment.CENTER,
                spacing=10
            ),
            ft.Container(
                content=self.scrollable_table,
                expand=True,
                border=ft.border.all(1, ft.Colors.GREY_400),
                border_radius=8,
                padding=10,
                alignment=ft.alignment.center,
            )
        ]

        # 设置 Column 的属性
        self.expand = True
        self.horizontal_alignment = ft.CrossAxisAlignment.STRETCH

        # 初始化表格数据
        if self.sheet_dropdown.value:
            self.update_table(self.sheet_dropdown.value)

    def update_table(self, sheet_name: str):
        if not sheet_name:
            self.data_table.columns = []
            self.data_table.rows = []
            self._page.update()
            return

        self.excel_tool.workbook = self.excel_tool.load_workbook(self.excel_tool.file_path) if os.path.exists(
            self.excel_tool.file_path) else self.excel_tool.create_workbook()
        sheet = self.excel_tool.get_sheet(sheet_name)
        row_count = self.excel_tool.get_row_count(sheet)
        col_count = self.excel_tool.get_column_count(sheet)

        headers = self.excel_tool.default_headers if self.excel_tool.check_headers(sheet) else [
            sheet.cell(row=self.excel_tool.header_row, column=col).value or f"Column {col}"
            for col in range(1, col_count + 1)
        ]

        self.data_table.columns = [
            ft.DataColumn(
                ft.Text(
                    header,
                    weight=ft.FontWeight.BOLD,
                    text_align=ft.TextAlign.CENTER,
                )
            )
            for header in headers
        ]

        if row_count > self.excel_tool.header_row:
            data = self.excel_tool.read_range(
                sheet, self.excel_tool.header_row + 1, 1, row_count, col_count, skip_headers=True
            )
            self.data_table.rows = [
                ft.DataRow(
                    cells=[
                        ft.DataCell(
                            ft.Text(
                                str(cell) if cell is not None else "",
                                text_align=ft.TextAlign.CENTER
                            ),
                        ) for cell in row
                    ]
                )
                for row in data
            ]
        else:
            self.data_table.rows = []

        self._page.update()

    def write_table_data(self, data):
        """
        写入表格数据
        :param data: 可以是 list 或 dict（包含 "values" 键）
        """
        sheet = self.excel_tool.get_sheet(self.excel_tool.get_sheet_names()[0])

        # 兼容两种数据格式
        if isinstance(data, dict):
            values = data.get("values", [])
        elif isinstance(data, list):
            values = data
        else:
            raise ValueError("data 必须是 list 或 dict 类型")

        # 跳过表头（第一行）
        for row_idx, row_data in enumerate(values[1:], start=self.excel_tool.header_row + 1):
            processed_row = []
            for col_idx, value in enumerate(row_data, start=1):
                if isinstance(value, list) and value and isinstance(value[0], dict) and "link" in value[0]:
                    cell_value = value[0]["link"] if value else None
                elif isinstance(value, (int, float)) and col_idx in [17, 18]:
                    try:
                        cell_value = datetime.fromordinal(
                            datetime(1900, 1, 1).toordinal() + int(value) - 2).strftime(
                            "%Y-%m-%d")
                    except (ValueError, TypeError):
                        cell_value = value
                else:
                    cell_value = value
                processed_row.append(cell_value)
                self.excel_tool.write_cell(
                    sheet=sheet,
                    row=row_idx,
                    col=col_idx,
                    value=cell_value,
                )
        self.excel_tool.save()
        self._page.update()

    def change_workbook(self, file_path: str):
        self.excel_tool.file_path = file_path
        self.excel_tool.workbook = self.excel_tool.load_workbook(file_path) if os.path.exists(
            file_path) else self.excel_tool.create_workbook()
        self.sheet_dropdown.options = [ft.dropdown.Option(sheet) for sheet in self.excel_tool.get_sheet_names()]
        self.sheet_dropdown.value = self.excel_tool.get_sheet_names()[0] if self.excel_tool.get_sheet_names() else None
        self.update_table(self.sheet_dropdown.value)

    def on_load_data_from_remote(self, e):
        main_app = self._page.data.get("main_app") if hasattr(self._page, "data") else None
        if not main_app:
            self._page.snack_bar = ft.SnackBar(ft.Text("无法找到主应用实例 ❌"))
            self._page.snack_bar.open = True
            self._page.update()
            print(f"无法找到主应用实例")
            return

        if not main_app.auth_status:
            self._page.snack_bar = ft.SnackBar(ft.Text("请先完成飞书授权 ✅"))
            self._page.snack_bar.open = True
            self._page.update()
            return

        token = main_app.token_storage.get("user_token")
        spreadsheets_token = main_app.sheet_storage.get("spreadsheets")
        sheet_id = main_app.sheet_storage.get("sheets")

        try:
            value_range = get_table_filter(token, spreadsheets_token, sheet_id)
            print(f"获取表数据成功:范围=> {value_range}")
            logger.success(f"获取表数据成功:范围=> {value_range}")
            values = get_table_value(token, spreadsheets_token, value_range)
            self.write_table_data(values)
            logger.success(f"获取表数据成功:=> {values[:10]}")
            self.update_table(self.sheet_dropdown.value)
        except Exception as ex:
            print(f"获取表数据失败: {ex}")
            logger.error(f"获取表数据失败: {ex}")
            self._page.update()

    def on_file_picked(self, e):
        if e.files and len(e.files) > 0:
            selected_file = e.files[0].path
            self.change_workbook(selected_file)
        else:
            self._page.snack_bar = ft.SnackBar(ft.Text("未选择文件"), open=True)
            self._page.update()

    def open_excel_file(self, e):
        if os.path.exists(self.excel_tool.file_path):
            os.startfile(self.excel_tool.file_path)
        else:
            self._page.snack_bar = ft.SnackBar(ft.Text("Excel 文件不存在"), open=True)
            self._page.update()