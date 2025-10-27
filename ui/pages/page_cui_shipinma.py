import flet as ft

class ShiPingMaPage(ft.Column):
    """查履约页面"""

    def __init__(self):
        super().__init__()
        self.expand = True
        self.horizontal_alignment = ft.CrossAxisAlignment.CENTER
        self.alignment = ft.MainAxisAlignment.CENTER

    def build(self):
        self.controls = [
        ]