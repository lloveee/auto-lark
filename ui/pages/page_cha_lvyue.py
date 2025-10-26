import flet as ft
class LvYuePage(ft.Column):
    """查履约页面"""

    def __init__(self, excel: ft.Control):
        super().__init__()
        self.expand = True
        self.horizontal_alignment = ft.CrossAxisAlignment.CENTER
        self.alignment = ft.MainAxisAlignment.START
        self.excel = excel

    def build(self):

        card_container = ft.Card(
            content=ft.Column(
                controls=[
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