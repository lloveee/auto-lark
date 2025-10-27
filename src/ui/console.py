"""
控制台 UI 组件
显示应用程序日志输出（支持拖拽调整高度）
"""
import flet as ft
from src.core.logger import logger


class Console(ft.Container):
    """控制台组件，可拖拽调整高度"""

    def __init__(self):
        super().__init__()
        self.log_container = None
        self.auto_scroll = True
        self.auto_scroll_btn = None
        self.expand = True
        self.height = 250  # 初始高度
        self.min_height = 120
        self.max_height = 600

    def build(self):
        """构建控制台界面"""
        # --- 日志显示区域 ---
        self.log_container = ft.Column(
            scroll=ft.ScrollMode.AUTO,
            auto_scroll=self.auto_scroll,
            spacing=2,
            expand=True,
        )

        # --- 控制按钮 ---
        clear_btn = ft.IconButton(
            icon=ft.Icons.DELETE_SWEEP,
            tooltip="清空日志",
            on_click=self._clear_logs,
            icon_size=20,
        )
        self.auto_scroll_btn = ft.IconButton(
            icon=ft.Icons.ARROW_DOWNWARD,
            tooltip="自动滚动",
            on_click=self._toggle_auto_scroll,
            icon_size=20,
            selected=True,
        )

        # --- 顶部栏 ---
        header = ft.Container(
            content=ft.Row(
                controls=[
                    ft.Container(expand=True),
                    self.auto_scroll_btn,
                    clear_btn,
                ],
            ),
            padding=ft.padding.only(left=10, right=10, top=0, bottom=0),
            bgcolor=ft.Colors.WHITE,
        )

        # --- 日志显示区域 ---
        log_area = ft.Container(
            content=self.log_container,
            padding=8,
            bgcolor=ft.Colors.WHITE,
            border_radius=ft.border_radius.only(bottom_left=5, bottom_right=5),
            expand=True,
        )

        # --- 拖拽分割条 ---
        def move_divider(e: ft.DragUpdateEvent):
            """在拖动时更新高度"""
            new_height = self.height - e.delta_y
            if self.min_height <= new_height <= self.max_height:
                self.height = new_height
                self.update()

        def show_resize_cursor(e: ft.HoverEvent):
            """当鼠标悬停在分割条上时显示调整光标"""
            e.control.mouse_cursor = ft.MouseCursor.RESIZE_UP_DOWN
            e.control.update()

        resize_bar = ft.GestureDetector(
            content=ft.Container(height=4, bgcolor=ft.Colors.GREY_300),
            on_pan_update=move_divider,
            on_hover=show_resize_cursor,
        )

        # --- 主体布局 ---
        self.content = ft.Column(
            controls=[resize_bar, log_area, header],
            spacing=0,
            horizontal_alignment=ft.CrossAxisAlignment.STRETCH,
        )

        self.border = ft.border.all(2, ft.Colors.OUTLINE_VARIANT)
        self.border_radius = 5

    # ---------------- 日志逻辑 ----------------
    def did_mount(self):
        logger.add_callback(self._on_log)
        self._load_history()

    def will_unmount(self):
        logger.remove_callback(self._on_log)

    def _load_history(self):
        logs = logger.get_logs()
        for log in logs[-50:]:
            self._add_log_line(log, update=False)
        if self.page:
            self.update()

    def _on_log(self, level: str, message: str):
        from datetime import datetime
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_line = f"[{timestamp}] [{level}] {message}"
        self._add_log_line(log_line)

    def _add_log_line(self, log_line: str, update=True):
        if self.log_container is None:
            return
        color = ft.Colors.ON_SURFACE
        if "[ERROR]" in log_line:
            color = ft.Colors.RED
        elif "[WARNING]" in log_line:
            color = ft.Colors.ORANGE_400
        elif "[SUCCESS]" in log_line:
            color = ft.Colors.GREEN_400
        elif "[INFO]" in log_line:
            color = ft.Colors.BLUE_400

        self.log_container.controls.append(
            ft.Text(
                log_line,
                size=14,
                color=color,
                font_family="Consolas",
                selectable=True,
                expand=True,
            )
        )

        if len(self.log_container.controls) > 100:
            self.log_container.controls.pop(0)
        if update and self.page:
            try:
                self.update()
            except Exception:
                pass

    def _clear_logs(self, e):
        if self.log_container:
            self.log_container.controls.clear()
            logger.clear()
            if self.page:
                self.update()

    def _toggle_auto_scroll(self, e):
        self.auto_scroll = not self.auto_scroll
        if self.log_container:
            self.log_container.auto_scroll = self.auto_scroll
            self.auto_scroll_btn.selected = self.auto_scroll
            if self.page:
                self.update()
