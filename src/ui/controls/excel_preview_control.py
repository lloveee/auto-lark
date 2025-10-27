from datetime import datetime
from src.modes.mode_cha_lvyue import get_table_filter
from src.modes.mode_cha_lvyue import get_table_value
from src.core.logger import logger
from src.modes.excel.excel_tool import ExcelTool
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
        self.apply_filter_button = None

        # 添加对话框和复选框存储
        self.sheet_selection_dialog = None
        self.sheet_checkboxes = {}

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
            on_click=self.show_sheet_selection_dialog,
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

    def show_sheet_selection_dialog(self, e):
        """显示表选择对话框"""
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

        # 获取所有表的字典 {title: sheet_id}
        sheets = main_app.sheet_storage.get("sheets", {})

        print(f"获取到的sheets: {sheets}")  # 调试信息

        if not sheets or not isinstance(sheets, dict):
            self._page.snack_bar = ft.SnackBar(ft.Text("没有可用的表 ❌"))
            self._page.snack_bar.open = True
            self._page.update()
            return

        # 清空之前的复选框
        self.sheet_checkboxes = {}
        checkbox_controls = []

        # 添加全选/取消全选按钮
        select_all_btn = ft.TextButton(
            "全选",
            on_click=lambda _: self.toggle_all_checkboxes(True)
        )
        deselect_all_btn = ft.TextButton(
            "取消全选",
            on_click=lambda _: self.toggle_all_checkboxes(False)
        )

        checkbox_controls.append(
            ft.Row([select_all_btn, deselect_all_btn], spacing=10)
        )
        checkbox_controls.append(ft.Divider())

        # 为每个表创建复选框，显示 title，但存储 sheet_id
        for title, sheet_id in sheets.items():
            checkbox = ft.Checkbox(
                label=title,  # 显示表格名称
                value=True,  # 默认全选
            )
            self.sheet_checkboxes[sheet_id] = checkbox  # 用 sheet_id 作为 key
            checkbox_controls.append(checkbox)

        print(f"创建了 {len(self.sheet_checkboxes)} 个复选框")  # 调试信息

        # 创建对话框
        self.sheet_selection_dialog = ft.AlertDialog(
            modal=True,
            title=ft.Text("选择要加载的表", size=20, weight=ft.FontWeight.BOLD),
            content=ft.Container(
                content=ft.Column(
                    controls=checkbox_controls,
                    scroll=ft.ScrollMode.AUTO,
                    tight=True,
                ),
                width=400,
                height=min(400, len(checkbox_controls) * 50 + 100),  # 动态高度
            ),
            actions=[
                ft.TextButton("取消", on_click=self.close_dialog),
                ft.ElevatedButton(
                    "确认加载",
                    on_click=self.confirm_load_data,  # 改为新的方法
                    style=ft.ButtonStyle(
                        bgcolor=ft.Colors.BLUE_400,
                        color=ft.Colors.WHITE,
                    )
                ),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
        )

        # 使用 page.open() 方法打开对话框
        self._page.open(self.sheet_selection_dialog)

        print("对话框已显示")  # 调试信息

    def toggle_all_checkboxes(self, value: bool):
        """全选或取消全选所有复选框"""
        for checkbox in self.sheet_checkboxes.values():
            checkbox.value = value
        self._page.update()

    def close_dialog(self, e):
        """关闭对话框"""
        if self.sheet_selection_dialog:
            self._page.close(self.sheet_selection_dialog)
            print("对话框已关闭")  # 调试信息

    def confirm_load_data(self, e):
        """确认加载数据"""
        # 获取选中的表
        selected_sheets = [
            sheet_id for sheet_id, checkbox in self.sheet_checkboxes.items()
            if checkbox.value
        ]

        print(f"选中的表: {selected_sheets}")  # 调试信息

        # 关闭对话框
        self.close_dialog(e)

        # 调用加载方法
        self.on_load_data_from_remote(selected_sheets)

    def on_load_data_from_remote(self, selected_sheets):
        """从远程加载数据（根据用户选择的表）"""
        main_app = self._page.data.get("main_app") if hasattr(self._page, "data") else None
        if not main_app:
            self._page.snack_bar = ft.SnackBar(ft.Text("无法找到主应用实例 ❌"))
            self._page.snack_bar.open = True
            self._page.update()
            print(f"无法找到主应用实例")
            return

        if not selected_sheets:
            self._page.snack_bar = ft.SnackBar(ft.Text("未选择任何表 ⚠️"))
            self._page.snack_bar.open = True
            self._page.update()
            return

        token = main_app.token_storage.get("user_token")
        spreadsheets_token = main_app.sheet_storage.get("spreadsheets")

        codelist = [["订单号"]]
        total_count = 0

        for sheet_id in selected_sheets:
            try:
                value_range = get_table_filter(token, spreadsheets_token, sheet_id)
                print(f"获取表数据成功:范围=> {value_range}")
                logger.success(f"获取表数据成功:范围=> {value_range}")

                values = get_table_value(token, spreadsheets_token, value_range)
                if not values or len(values) < 2:
                    logger.warning(f"{sheet_id} 表数据为空或无效")
                    continue
                header = values[0]
                # 必须同时存在"序号"和"履约方式"
                if "序号" not in header or "履约方式" not in header or "订单号" not in header:
                    logger.warning(f"{sheet_id} 表头不含关键列，跳过：{header}")
                    continue

                idx_xh = header.index("序号")
                idx_ly = header.index("履约方式")
                idx_order = header.index("订单号")
                i = 0
                for row in values[1:]:
                    try:
                        xh = row[idx_xh] if idx_xh < len(row) else None
                        ly = row[idx_ly] if idx_ly < len(row) else None
                        order = row[idx_order] if idx_order < len(row) else None

                        if xh not in (None, "", "null") and ly in (None, "", "null") and order not in (
                        None, "", "null"):
                            codelist.append([order])
                            i = i + 1
                    except Exception as row_ex:
                        logger.warning(f"解析行时出错：{row_ex} => {row}")

                total_count += i
                logger.success(f"{sheet_id} ✅筛选完成，共提取 {i} 条")

            except Exception as ex:
                print(f"获取表数据失败: {ex}")
                logger.error(f"获取表数据失败: {ex}")

        try:
            self.write_table_data(codelist)
            self.update_table(self.sheet_dropdown.value)
            self._page.snack_bar = ft.SnackBar(
                ft.Text(f"✅ 成功加载 {total_count} 条数据"),
                open=True
            )
            self._page.update()
        except Exception as ex:
            logger.error(f"提取订单号失败=>{ex}")
            self._page.snack_bar = ft.SnackBar(
                ft.Text(f"❌ 加载失败: {ex}"),
                open=True
            )
            self._page.update()

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
        完全覆写 Excel 表格，包括表头
        :param data: 可以是 list 或 dict（包含 "values" 键）
        """
        sheet_name = self.excel_tool.get_sheet_names()[0]
        sheet = self.excel_tool.get_sheet(sheet_name)

        # 清空整个表格
        max_row = sheet.max_row
        max_col = sheet.max_column
        if max_row > 0:
            sheet.delete_rows(1, max_row)
        if max_col > 0:
            sheet.delete_cols(1, max_col)

        # 兼容两种数据格式
        if isinstance(data, dict):
            values = data.get("values", [])
        elif isinstance(data, list):
            values = data
        else:
            raise ValueError("data 必须是 list 或 dict 类型")

        if not values:
            self.excel_tool.save()
            self._page.update()
            return

        # 写入表头 + 数据
        for row_idx, row_data in enumerate(values, start=1):
            for col_idx, value in enumerate(row_data, start=1):
                # 处理链接类型
                if isinstance(value, list) and value and isinstance(value[0], dict) and "link" in value[0]:
                    cell_value = value[0]["link"]
                # 处理日期列（17、18列）
                elif isinstance(value, (int, float)) and col_idx in [17, 18]:
                    try:
                        cell_value = datetime.fromordinal(
                            datetime(1900, 1, 1).toordinal() + int(value) - 2
                        ).strftime("%Y-%m-%d")
                    except (ValueError, TypeError):
                        cell_value = value
                else:
                    cell_value = value

                self.excel_tool.write_cell(
                    sheet=sheet,
                    row=row_idx,
                    col=col_idx,
                    value=cell_value,
                )

        self.excel_tool.save()
        self._page.update()

    def filter_non_empty_rows(self, selected_cols: list[str]):
        """剔除指定列为空的行，并覆写 Excel 文件"""
        file_path = self.excel_tool.file_path
        if not os.path.exists(file_path):
            self._page.snack_bar = ft.SnackBar(ft.Text("文件不存在 ❌"), open=True)
            self._page.update()
            return

        try:
            sheet_name = self.sheet_dropdown.value
            sheet = self.excel_tool.get_sheet(sheet_name)
            row_count = self.excel_tool.get_row_count(sheet)
            col_count = self.excel_tool.get_column_count(sheet)

            headers = [
                sheet.cell(row=self.excel_tool.header_row, column=col).value or f"Column {col}"
                for col in range(1, col_count + 1)
            ]

            selected_col_indices = [
                headers.index(col_name) + 1 for col_name in selected_cols if col_name in headers
            ]

            data_rows = self.excel_tool.read_range(
                sheet,
                start_row=self.excel_tool.header_row + 1,
                start_col=1,
                end_row=row_count,
                end_col=col_count,
                skip_headers=True
            )

            filtered_rows = []
            for row in data_rows:
                keep = all(
                    str(row[idx - 1]).strip() != "" and row[idx - 1] is not None
                    for idx in selected_col_indices
                )
                if keep:
                    filtered_rows.append(row)

            # 清空旧数据
            for r in range(self.excel_tool.header_row + 1, row_count + 1):
                for c in range(1, col_count + 1):
                    sheet.cell(row=r, column=c).value = None

            # 写回过滤后的数据
            self.excel_tool.write_range(sheet, self.excel_tool.header_row + 1, 1, filtered_rows)
            self.excel_tool.save()
            self.update_table(sheet_name)

            msg = f"✅ 已剔除 {len(data_rows) - len(filtered_rows)} 行空值数据，并覆写文件"
            logger.success(msg)
            self._page.snack_bar = ft.SnackBar(ft.Text(msg), open=True)
            self._page.update()

        except Exception as ex:
            logger.error(f"剔除空值失败: {ex}")
            self._page.snack_bar = ft.SnackBar(ft.Text(f"剔除失败: {ex}"), open=True)
            self._page.update()

    def change_workbook(self, file_path: str):
        try:
            old = self.excel_tool.file_path
            self.excel_tool.file_path = file_path
            logger.success(f"切换文件路径{old} => {file_path}")
            self.excel_tool.workbook = self.excel_tool.load_workbook(file_path) if os.path.exists(
                file_path) else self.excel_tool.create_workbook()
            self.sheet_dropdown.options = [ft.dropdown.Option(sheet) for sheet in self.excel_tool.get_sheet_names()]
            self.sheet_dropdown.value = self.excel_tool.get_sheet_names()[
                0] if self.excel_tool.get_sheet_names() else None
            # self.update_table(self.sheet_dropdown.value)
        except Exception as e:
            logger.error(f"加载文件失败：{e}")

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