import flet as ft
from src.core.logger import logger

class BrowserTab(ft.Column):
    """单个浏览器标签页"""

    def __init__(self, tab_id: int, on_close=None):
        super().__init__()
        self.tab_id = tab_id
        self.on_close = on_close
        self.url_input = None
        self.web_view = None
        self.current_url = "https://flet.dev"
        self.expand = True
        self.spacing = 0

    def build(self):
        """构建浏览器页面"""
        self.url_input = ft.TextField(
            value=self.current_url,
            hint_text="输入网址...",
            prefix_icon=ft.Icons.SEARCH,
            expand=True,
            on_submit=self._navigate,
            border_radius=20,
            height=45,
        )

        back_btn = ft.IconButton(
            icon=ft.Icons.ARROW_BACK,
            tooltip="后退",
            on_click=self._go_back,
        )

        forward_btn = ft.IconButton(
            icon=ft.Icons.ARROW_FORWARD,
            tooltip="前进",
            on_click=self._go_forward,
        )

        refresh_btn = ft.IconButton(
            icon=ft.Icons.REFRESH,
            tooltip="刷新",
            on_click=self._refresh,
        )

        home_btn = ft.IconButton(
            icon=ft.Icons.HOME,
            tooltip="主页",
            on_click=self._go_home,
        )

        # 导航栏
        navbar = ft.Container(
            content=ft.Row(
                controls=[
                    back_btn,
                    forward_btn,
                    refresh_btn,
                    home_btn,
                    ft.VerticalDivider(width=1),
                    self.url_input,
                ],
                spacing=5,
            ),
            padding=10,
            bgcolor=ft.Colors.GREY_50,
            border_radius=10,
        )

        # 使用 HTML + IFrame 方式显示网页 (Windows 桌面端兼容方案)
        self.web_view = ft.WebView(
            url=self.current_url,
            expand=True,
            on_page_started=lambda e: self._on_page_started(e),
            on_page_ended=lambda e: self._on_page_ended(e),
            on_web_resource_error=lambda e: logger.error(f"标签页 {self.tab_id} 加载错误: {e.data}"),
        )

        # 设置Column属性
        self.controls = [
            navbar,
            ft.Container(height=10),
            ft.Container(
                content=self.web_view,
                border=ft.border.all(1, ft.Colors.OUTLINE_VARIANT),
                border_radius=10,
                expand=True,
            ),
        ]

    def _navigate(self, e):
        """导航到URL"""
        url = self.url_input.value.strip()
        if url:
            if not url.startswith(("http://", "https://")):
                url = "https://" + url
            self.current_url = url

            # 更新显示的URL
            if self.web_view and self.web_view.content:
                url_text = self.web_view.content.controls[1]
                url_text.value = f"目标URL: {url}"
                self.web_view.update()

            logger.info(f"标签页 {self.tab_id} URL已更新: {url}")

    def _open_in_browser(self, e):
        """在系统浏览器中打开"""
        import webbrowser
        try:
            webbrowser.open(self.current_url)
            logger.success(f"已在浏览器中打开: {self.current_url}")
        except Exception as ex:
            logger.error(f"打开浏览器失败: {ex}")

    def _refresh(self, e):
        """刷新页面"""
        logger.info(f"标签页 {self.tab_id} 刷新 (功能仅在移动端可用)")

    def _go_back(self, e):
        """后退"""
        logger.info(f"标签页 {self.tab_id} 后退 (功能仅在移动端可用)")

    def _go_forward(self, e):
        """前进"""
        logger.info(f"标签页 {self.tab_id} 前进 (功能仅在移动端可用)")

    def _go_home(self, e):
        """回到主页"""
        self.url_input.value = "https://flet.dev"
        self._navigate(e)

    def _on_page_started(self, e):
        """页面开始加载"""
        logger.info(f"标签页 {self.tab_id} 开始加载...")
        if self.url_input and hasattr(e, 'data'):
            self.url_input.value = e.data
            self.url_input.update()

    def _on_page_ended(self, e):
        """页面加载完成"""
        logger.success(f"标签页 {self.tab_id} 加载完成")