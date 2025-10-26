import flet as ft
import flet_webview as ftwv
from core.logger import logger
from core.env import SPREADSHEET_TOKEN
from modes.mode_drive_api import get_spreadsheetToken, get_spreadsheet_Id

class HomePage(ft.Column):
    """主页标签页"""

    def __init__(self):
        super().__init__()
        self.expand = True
        self.horizontal_alignment = ft.CrossAxisAlignment.START
        self.dropdown_spreadsheet = None
        self.dropdown_sheet = None
        self.files = []
        self.has_cache = False


    def build(self):
        # 获取 MainApp 实例（父容器）
        main_app = self.page.data.get("main_app") if hasattr(self, "page") else None
        if not main_app:
            self.page.snack_bar = ft.SnackBar(ft.Text("无法找到主应用实例 ❌"))
            self.page.snack_bar.open = True
            self.page.update()
            return

        spreadsheets = main_app.sheet_storage.get("spreadsheets")
        sheets = main_app.sheet_storage.get("sheets")

        if spreadsheets and sheets:
            self.has_cache = True

        spreadsheet_options = (
            [ft.dropdown.Option(key) for key in spreadsheets]
            if spreadsheets and isinstance(spreadsheets, (dict, list)) and len(spreadsheets) > 0
            else []
        )

        hint_text = "暂无数据" if not spreadsheet_options else "选择表格token"

        self.dropdown_spreadsheet = ft.Dropdown(
            value=spreadsheet_options[0] if spreadsheet_options and len(spreadsheet_options) > 0 else None,
            label=hint_text,
            options=spreadsheet_options,
            width=200,
            disabled=not spreadsheet_options,
            on_change=self._on_spread_sheet_change,
        )

        sheets_options = (
            [ft.dropdown.Option(key) for key in sheets]
            if sheets and isinstance(sheets, (dict, list)) and len(sheets) > 0
            else []
        )

        hint_text = "暂无数据" if not sheets_options else "选择表ID"

        self.dropdown_sheet = ft.Dropdown(
            value= sheets_options[0] if sheets_options and len(sheets_options) > 0 else None,
            label=hint_text,
            options=sheets_options,
            width=200,
            disabled=not sheets_options,
            on_change=self._on_sheet_change,
        )

        update_button = ft.ElevatedButton(
            text="重新获取数据",
            icon=ft.Icons.REFRESH,
            on_click=self.on_update_cache_click,
        )

        # 配置板块容器
        config_section = ft.Container(
            content=ft.Column(
                controls=[
                    ft.Text("全局数据配置", size=20, weight=ft.FontWeight.BOLD),
                    ft.Row(
                        controls=[
                            self.dropdown_spreadsheet,
                            self.dropdown_sheet,
                            update_button,
                        ],
                        spacing=20,
                        alignment=ft.MainAxisAlignment.CENTER,
                    ),

                ],
                spacing=10,
                horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            ),
            padding=ft.padding.all(15),
            bgcolor=ft.Colors.WHITE,
            margin=ft.margin.only(bottom=15, top=10),
            shadow=ft.BoxShadow(
                spread_radius=0.5,
                blur_radius=4,
                color=ft.Colors.GREY_300,
                offset=ft.Offset(0, 2),
            ),
            border_radius=10,
        )

        """构建主页"""
        self.controls = [
            config_section,
        ]

    def _on_type_change(self, e):
        self._update_spreadsheet_dropdown()

    def _update_spreadsheet_dropdown(self):
        filtered_files = [f for f in self.files if f['type'] == 'sheet']

        options = [
            ft.dropdown.Option(key=f['token'], text=f['name']) for f in filtered_files
        ]
        logger.info(f"sheet_token=>{options[0].key}" )

        self.dropdown_spreadsheet.options = options
        self.dropdown_spreadsheet.value = options[0].key if options else None
        self.dropdown_spreadsheet.disabled = len(options) == 0
        self.dropdown_spreadsheet.label = "暂无数据" if len(options) == 0 else "选择表格"

        # 如果有选中值，触发更新 sheet 下拉框
        if self.dropdown_spreadsheet.value:
            self._on_spread_sheet_change(None)
        else:
            # 清空 sheet 下拉框
            self.dropdown_sheet.options = []
            self.dropdown_sheet.value = None
            self.dropdown_sheet.disabled = True

        self.page.update()

    def _on_spread_sheet_change(self, e):
        main_app = self.page.data.get("main_app") if hasattr(self, "page") else None
        if main_app and main_app.auth_status:
            spreadsheet_token = self.dropdown_spreadsheet.value
            main_app.sheet_storage.set("spreadsheets", self.dropdown_spreadsheet.value)
            logger.info(f"表格token更新为{self.dropdown_spreadsheet.value}")
            access_token = main_app.token_storage.get("user_token")

            try:
                sheets = get_spreadsheet_Id(access_token, spreadsheet_token)
                options = [
                    ft.dropdown.Option(key=s['sheet_id'], text=s['title']) for s in sheets
                ]
                self.dropdown_sheet.options = options
                self.dropdown_sheet.value = options[0].key if options else None
                self.dropdown_sheet.disabled = len(options) == 0
                self.dropdown_sheet.label = "暂无数据" if len(options) == 0 else "选择表"

                # 如果有选中值，触发 sheet change
                if self.dropdown_sheet.value:
                    self._on_sheet_change(None)
            except Exception as ex:
                logger.error(f"获取 sheets 失败: {ex}")
                self.page.snack_bar = ft.SnackBar(ft.Text(f"获取 sheets 失败: {str(ex)}"))
                self.page.snack_bar.open = True
            self.page.update()
        else:
            logger.error("请先完成飞书授权")
    def _on_sheet_change(self, e):
        main_app = self.page.data.get("main_app") if hasattr(self, "page") else None
        if main_app and main_app.auth_status:
            main_app.sheet_storage.set("sheets", self.dropdown_sheet.value)
            logger.info(f"表id更新为{self.dropdown_sheet.value}")
        else:
            logger.error("请先完成飞书授权")

    def on_update_cache_click(self, e):
        main_app = self.page.data.get("main_app") if hasattr(self, "page") else None
        if main_app and main_app.auth_status:
            access_token = main_app.token_storage.get("user_token")
            try:
                self.files = get_spreadsheetToken(access_token)
                logger.info(f"获取文档数据成功=>{self.files}")
                self._update_spreadsheet_dropdown()
                self.page.snack_bar = ft.SnackBar(ft.Text("数据更新成功 ✅"))
                self.page.snack_bar.open = True
            except Exception as ex:
                logger.error(f"获取 files 失败: {ex}")
                self.page.snack_bar = ft.SnackBar(ft.Text(f"获取 files 失败: {str(ex)}"))
                self.page.snack_bar.open = True

            self.page.update()
        else:
            logger.error("请先进行飞书授权")

    def _add_browser_tab(self, e):
        """触发添加浏览器标签"""
        # 通过page的data传递命令
        if self.page and self.page.data and 'main_app' in self.page.data:
            self.page.data['main_app'].add_new_browser_tab(None)