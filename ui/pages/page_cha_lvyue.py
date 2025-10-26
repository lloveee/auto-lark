import flet as ft
from modes.excel.excel_tool import ExcelTool
from modes.mode_cha_lvyue import get_table_filter
from core.logger import logger
class LvYuePage(ft.Column):
    """查履约页面"""

    def __init__(self, excel: ft.Control):
        super().__init__()
        self.expand = True
        self.horizontal_alignment = ft.CrossAxisAlignment.CENTER
        self.alignment = ft.MainAxisAlignment.START
        self.excel = excel

        # 输入框
        self.filter_range = ft.TextField(label="筛选范围", hint_text="示例: A1:D50", width=160)
        self.filter_col = ft.TextField(label="列号", hint_text="示例: 2 或 B", width=120)
        self.filter_type = ft.TextField(label="filter_type", hint_text="示例: VALUE", width=160)
        self.compare_type = ft.TextField(label="compare_type", hint_text="示例: EQUAL", width=160)
        self.expected = ft.TextField(label="expected", hint_text="示例: 已完成", width=160)

        self.run_button = ft.ElevatedButton(
            text="执行筛选",
            icon=ft.Icons.SEARCH,
            style=ft.ButtonStyle(
                bgcolor=ft.Colors.BLUE_400,
                color=ft.Colors.WHITE,
                shape=ft.RoundedRectangleBorder(radius=10),
                padding=ft.padding.symmetric(horizontal=20, vertical=10),
            ),
            on_click=self._on_run_filter_click,
        )

    def build(self):
        example_text = ft.Text(
            "举例：范围 A1:D50，列号 2，filter_type=VALUE，compare_type=EQUAL，expected=已完成",
            size=12,
            color=ft.Colors.GREY_600,
            italic=True,
        )
        filter_row = ft.Row(
            controls=[
                self.filter_range,
                self.filter_col,
                self.filter_type,
                self.compare_type,
                self.expected,
                self.run_button,
            ],
            alignment=ft.MainAxisAlignment.START,
            vertical_alignment=ft.CrossAxisAlignment.CENTER,
            spacing=10,
            wrap=True,  # 当窗口较窄时自动换行
        )

        filter_container = ft.ExpansionPanelList(
            controls=[
                ft.ExpansionPanel(
                    header=ft.ListTile(ft.Text("筛选条件", size=14, weight=ft.FontWeight.BOLD, color=ft.Colors.BLUE_600)),
                    content=ft.Container(
                        padding=ft.padding.only(bottom=15),  # Add padding at the bottom
                        content=ft.Column(
                            controls=[
                                example_text,
                                ft.Divider(height=5, color=ft.Colors.GREY_200),
                                filter_row,
                            ],
                            spacing=8,
                            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                        ),
                    ),
                    bgcolor=ft.Colors.GREY_100,
                    expanded=False,
                )
            ],
            expand_icon_color=ft.Colors.BLUE_600,
            divider_color=ft.Colors.AMBER,
            elevation=8,
        )

        card_container = ft.Card(
            content=ft.Column(
                controls=[
                    filter_container,
                    self.excel,
                ],
                spacing=10,
                horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            ),
            expand=True,
            elevation=4,
            margin=ft.margin.all(10),
            color=ft.Colors.WHITE,
            shadow_color=ft.Colors.GREY_400,
            surface_tint_color=ft.Colors.GREY_100,
        )

        # 页面控件集合
        self.controls = [
            card_container,
        ]

    def _on_run_filter_click(self, e):
        """点击筛选按钮时执行"""
        # 获取 MainApp 实例（父容器）
        main_app = self.page.data.get("main_app") if hasattr(self, "page") else None

        if not main_app:
            self.page.snack_bar = ft.SnackBar(ft.Text("无法找到主应用实例 ❌"))
            self.page.snack_bar.open = True
            self.page.update()
            return

        if not main_app.auth_status:
            self.page.snack_bar = ft.SnackBar(ft.Text("请先完成飞书授权 ✅"))
            self.page.snack_bar.open = True
            self.page.update()
            return

        token = main_app.token_storage.get("user_token")
        if not token:
            self.page.snack_bar = ft.SnackBar(ft.Text("未找到有效Token，请重新授权 ❌"))
            self.page.snack_bar.open = True
            self.page.update()
            return

        # 读取输入参数
        filter_range = self.filter_range.value.strip()
        filter_col = self.filter_col.value.strip()
        filter_type = self.filter_type.value.strip()
        compare_type = self.compare_type.value.strip()
        expected = self.expected.value.strip()

        if not all([filter_range, filter_col, filter_type, compare_type, expected]):
            self.page.snack_bar = ft.SnackBar(ft.Text("请填写所有筛选参数 ⚠️"))
            self.page.snack_bar.open = True
            self.page.update()
            return

        try:
            result = get_table_filter(token, filter_range, filter_col, filter_type, compare_type, expected)
            self.page.snack_bar = ft.SnackBar(ft.Text("筛选成功 ✅"))
            self.page.snack_bar.open = True
            self.page.update()
            logger.success(f"筛选成功: {result}")
        except Exception as ex:
            logger.error(f"筛选失败: {ex}")
            self.page.snack_bar = ft.SnackBar(ft.Text(f"筛选失败: {ex}"))
            self.page.snack_bar.open = True
            self.page.update()