import flet as ft
import flet_webview as ftwv
from core.logger import logger

class HomePage(ft.Column):
    """主页标签页"""

    def __init__(self):
        super().__init__()
        self.expand = True
        self.horizontal_alignment = ft.CrossAxisAlignment.START
        self.dropdown_spreadsheet = None
        self.dropdown_sheet = None

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
            on_click=self._on_update_cache_click,
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

    def _on_spread_sheet_change(self, e):
        main_app = self.page.data.get("main_app") if hasattr(self, "page") else None
        if main_app:
            main_app.drive_data["spreadsheet"] = self.dropdown_spreadsheet.value
            logger.info(f"表格token更新为{self.dropdown_spreadsheet.value}")
    def _on_sheet_change(self, e):
        main_app = self.page.data.get("main_app") if hasattr(self, "page") else None
        if main_app:
            main_app.drive_data["sheet"] = self.dropdown_sheet.value
            logger.info(f"表id更新为{self.dropdown_sheet.value}")

    def _on_update_cache_click(self, e):
        pass

    def _add_browser_tab(self, e):
        """触发添加浏览器标签"""
        # 通过page的data传递命令
        if self.page and self.page.data and 'main_app' in self.page.data:
            self.page.data['main_app'].add_new_browser_tab(None)