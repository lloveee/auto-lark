"""
主应用程序入口
实现类似浏览器的多标签页框架
"""
import flet as ft
import flet_webview as ftwv
import asyncio
import time
from ui.console import Console
from core.logger import logger
from ui.pages.home_page import HomePage
from ui.pages.browser_tab import BrowserTab
from ui.pages.page_cha_lvyue import LvYuePage
from ui.pages.page_tongguo_yangpin import YangPingPage
from ui.pages.page_cui_shipinma import ShiPingMaPage
from ui.pages.page_cha_daohuo import DaoHuoPage
from modes.feishu.feishu_auth import start_authorize_flow, exchange_code_for_token
from modes.feishu.feishu_sheet import append_to_sheet, col_num_to_letter
from core.env import TOKEN_STORE_FILE, SHEET_STORE_FILE
from modes.persistence.storage import Storage

class MainApp:
    """主应用程序"""

    def __init__(self, page: ft.Page):

        self.page = page
        self.access_token = None
        self.tab_counter = 0
        self.tabs_control = None
        self.console_container = None
        self.console_expanded = True
        self.auth_status = False
        self.auth_indicator = None
        #storage
        self.token_storage = Storage(TOKEN_STORE_FILE)
        self.sheet_storage = Storage(SHEET_STORE_FILE)
        # 保存引用到page.data供其他组件访问
        if self.page.data is None:
            self.page.data = {}
        self.page.data['main_app'] = self
        self.setup_page()
        self.build_ui()
        self._check_stored_token()
        self.home_page.on_update_cache_click(None)
        # 启动异步服务器
        self.page.run_task(self._start_local_server)
        logger.success("回调服务器启动 => 127.0.0.1:3000/callback")
    def setup_page(self):
        """配置页面"""
        self.page.title = "现代化飞书框架"
        self.page.theme_mode = ft.ThemeMode.LIGHT
        self.page.padding = 0
        self.page.window.width = 1200
        self.page.window.height = 800
        self.page.window.min_width = 900
        self.page.window.min_height = 600

    def build_ui(self):
        """构建用户界面"""
        # 授权状态指示器
        self.auth_indicator = ft.Container(
            content=ft.Row(
                controls=[
                    ft.Icon(
                        ft.Icons.LOCK,
                        size=16,
                        color=ft.Colors.RED_400,
                    ),
                    ft.Text(
                        "飞书用户未授权",
                        size=12,
                        weight=ft.FontWeight.BOLD,
                        color=ft.Colors.RED_400,
                    ),
                ],
                spacing=5,
                tight=True,
            ),
            padding=ft.padding.symmetric(horizontal=12, vertical=6),
            bgcolor=ft.Colors.GREY_50,
            border_radius=15,
            border=ft.border.all(1, ft.Colors.RED_400),
            tooltip="授权状态(只读)",
        )

        #self._set_auth_status()

        # 尝试调用飞书授权
        auth_button = ft.IconButton(
            icon=ft.Icons.KEY,
            tooltip="授权飞书账号",
            on_click=self._toggle_auth,
            icon_size=20,
        )
        # 添加浏览器标签页按钮
        add_tab_btn = ft.IconButton(
            icon=ft.Icons.ADD,
            tooltip="新建浏览器标签页",
            on_click=None,#self.add_new_browser_tab,
            icon_size=20,
            style=ft.ButtonStyle(
                shape=ft.CircleBorder(),
            ),
        )

        # 头部栏
        header = ft.Container(
            content=ft.Row(
                controls=[
                    ft.Icon(ft.Icons.LANGUAGE, size=24, color=ft.Colors.PRIMARY),
                    ft.Text(
                        "现代化飞书工具",
                        size=16,
                        weight=ft.FontWeight.BOLD,
                    ),
                    ft.Container(expand=True),
                    self.auth_indicator,
                    auth_button,
                    ft.VerticalDivider(width=10),
                    add_tab_btn,
                ],
                alignment=ft.MainAxisAlignment.START,
                vertical_alignment=ft.CrossAxisAlignment.CENTER,
            ),
            padding=ft.padding.symmetric(horizontal=15, vertical=10),
            bgcolor=ft.Colors.GREY_50,
        )

        # 初始化标签页控件
        self.tabs_control = ft.Tabs(
            selected_index=0,
            animation_duration=200,
            scrollable=True,
            tabs=[],
            expand=True,
            on_change=self._on_tab_change,
        )

        self._add_home_tab(0)
        self._add_lvyue_tab(1, "履约")
        self._add_yangping_tab(2, "样品")
        self._add_shiping_tab(3, "视频码")
        self._add_daohuo_tab(4, "到货")

        # 标签页容器
        tabs_container = ft.Container(
            content=self.tabs_control,
            padding=10,
            expand=True,
        )
        # 控制台
        console = Console()

        # 控制台控制按钮
        console_toggle_btn = ft.IconButton(
            icon=ft.Icons.KEYBOARD_ARROW_DOWN,
            tooltip="最小化控制台",
            on_click=self._toggle_console,
            icon_size=20,
        )

        console_header = ft.Container(
            content=ft.Row(
                controls=[
                    ft.Icon(ft.Icons.TERMINAL, size=18, color=ft.Colors.PRIMARY),
                    ft.Text("JJShell    版权所有(C) CKLJJ", size=14, weight=ft.FontWeight.BOLD),
                    ft.Container(expand=True),
                    console_toggle_btn,
                ],
                spacing=10,
            ),
            padding=ft.padding.symmetric(horizontal=10, vertical=5),
            bgcolor=ft.Colors.GREY_200,
            on_click=lambda e: self._toggle_console(e) if not self.console_expanded else None,
        )

        self.console_container = ft.Container(
            content=ft.Column(
                controls=[
                    console_header,
                    console,
                ],
                spacing=0,
            ),
            border=ft.border.all(1, ft.Colors.OUTLINE_VARIANT),
            border_radius=5,
            animate=ft.Animation(200, ft.AnimationCurve.EASE_IN_OUT),
        )
        # 主布局
        main_content = ft.Column(
            controls=[
                header,
                tabs_container,
                ft.Container(
                    content=self.console_container,
                    padding=ft.padding.only(left=10, right=10, bottom=10),
                ),
            ],
            spacing=0,
            expand=True,
        )

        self.page.add(
            ft.Container(
                content=main_content,
                bgcolor=ft.Colors.WHITE,
                expand=True,
            )
        )
        # 欢迎日志
        logger.info("应用程序启动成功")
        logger.success("欢迎使用CKLJJ现代化飞书")

    def _check_stored_token(self):
        """检查存储的令牌是否有效"""
        token = self.token_storage.get("user_token")
        expire_time = self.token_storage.get("expire_time")
        if token and expire_time:
            current_time = int(time.time())
            if current_time < expire_time:
                logger.info(f"找到有效令牌: {token[:10]}...，到期时间: {expire_time}")
                self.access_token = token
                self._set_auth_status(True)
                return True
            else:
                logger.info("存储的令牌已过期，需重新授权")
                self.token_storage.delete("user_token")
                self.token_storage.delete("expire_time")
        else:
            logger.info("未找到有效令牌，需进行授权")
        return False

    def _add_home_tab(self, index):
        """添加主页标签"""
        self.home_page = HomePage()

        # 创建主页标签
        tab = ft.Tab(
            content=self.home_page,
            tab_content=ft.Row(
                controls=[
                    ft.Icon(ft.Icons.HOME, size=16),
                    ft.Text("主页", size=13),
                ],
                spacing=5,
            ),
        )

        # 保存tab信息
        tab.data = {"tab_id": index, "tab_type": "home"}

        self.tabs_control.tabs.append(tab)
        self.page.update()

        logger.info("主页已加载")

    def _add_lvyue_tab(self, index, text):
        from modes.excel.excel_tool import ExcelTool
        from ui.controls.excel_preview_control import ExcelPreviewControl
        # 添加标签
        lvyue_excel = ExcelTool(
            file_name="查履约test.xlsx",
            header_row=1,
        )

        lvyue_excel.save()
        excel_panel = ExcelPreviewControl(lvyue_excel, self.page)
        lvyue_page = LvYuePage(excel_panel)
        tab = ft.Tab(
            content=lvyue_page,
            tab_content=ft.Row(
                controls=[
                    ft.Icon(ft.Icons.RULE, size=16),
                    ft.Text(text, size=13),
                ],
                spacing=5,
            ),
        )
        tab.data = {"tab_id": index, "tab_type": "lvyue"}
        self.tabs_control.tabs.append(tab)
        self.page.update()
        logger.info(f"{text}页面已加载")
    def _add_yangping_tab(self, index, text):
        yangping_page = YangPingPage()

        tab = ft.Tab(
            content=yangping_page,
            tab_content=ft.Row(
                controls=[
                    ft.Icon(ft.Icons.GIF_BOX, size=16),
                    ft.Text(text, size=13),
                ],
                spacing=5,
            ),
        )

        tab.data = {"tab_id": index, "tab_type": "yangping"}

        self.tabs_control.tabs.append(tab)
        self.page.update()

        logger.info(f"{text}页面已加载")
    def _add_shiping_tab(self, index, text):
        shipingma_page = ShiPingMaPage()

        tab = ft.Tab(
            content=shipingma_page,
            tab_content=ft.Row(
                controls=[
                    ft.Icon(ft.Icons.CODE, size=16),
                    ft.Text(text, size=13),
                ],
                spacing=5,
            ),
        )

        tab.data = {"tab_id": index, "tab_type": "shipingma"}

        self.tabs_control.tabs.append(tab)
        self.page.update()

        logger.info(f"{text}页面已加载")

    def _add_daohuo_tab(self, index, text):
        daohuo_page = DaoHuoPage()

        tab = ft.Tab(
            content=daohuo_page,
            tab_content=ft.Row(
                controls=[
                    ft.Icon(ft.Icons.ACCESSIBLE, size=16),
                    ft.Text(text, size=13),
                ],
                spacing=5,
            ),
        )

        tab.data = {"tab_id": index, "tab_type": "daohuo"}

        self.tabs_control.tabs.append(tab)
        self.page.update()

        logger.info(f"{text}页面已加载")


    def add_new_browser_tab(self, e):
        """添加新浏览器标签页"""
        self.tab_counter += 1
        tab_id = self.tab_counter

        # 创建新标签页内容
        browser_tab = BrowserTab(
            tab_id=tab_id,
            on_close=lambda: self._close_tab(tab_id)
        )

        # 关闭按钮
        close_btn = ft.IconButton(
            icon=ft.Icons.CLOSE,
            icon_size=16,
            on_click=lambda e: self._close_tab(tab_id),
            tooltip="关闭标签页",
        )

        # 创建标签
        tab = ft.Tab(
            content=browser_tab,
            tab_content=ft.Row(
                controls=[
                    ft.Icon(ft.Icons.WEB, size=16),
                    ft.Text(f"浏览器 {tab_id}", size=13),
                    close_btn,
                ],
                spacing=5,
            ),
        )

        # 保存tab信息
        tab.data = {"tab_id": tab_id, "tab_type": "browser", "browser_tab": browser_tab}

        self.tabs_control.tabs.append(tab)
        self.tabs_control.selected_index = len(self.tabs_control.tabs) - 1
        self.page.update()

        logger.info(f"新建浏览器标签页 {tab_id}")
    def _close_tab(self, tab_id: int):
        """关闭标签页"""
        # 找到并移除标签页 (不能关闭主页)
        for i, tab in enumerate(self.tabs_control.tabs):
            if hasattr(tab, 'data') and tab.data:
                if tab.data.get("tab_type") == "browser" and tab.data.get("tab_id") == tab_id:
                    self.tabs_control.tabs.pop(i)
                    if self.tabs_control.selected_index >= len(self.tabs_control.tabs):
                        self.tabs_control.selected_index = len(self.tabs_control.tabs) - 1
                    self.page.update()
                    logger.info(f"关闭浏览器标签页 {tab_id}")
                    break

    def _on_tab_change(self, e):
        """标签页切换事件"""
        current_tab = self.tabs_control.tabs[e.control.selected_index]
        if hasattr(current_tab, 'data') and current_tab.data:
            tab_type = current_tab.data.get("tab_type", "unknown")
            if tab_type == "home":
                logger.info("切换到主页")
            elif tab_type == "browser":
                logger.info(f"切换到浏览器标签页 {current_tab.data.get('tab_id')}")
            elif tab_type == "lvyue":
                logger.info("切换到履约页面")
            elif tab_type == "yangping":
                logger.info("切换到样品页面")
    
    def _toggle_console(self, e):
        """切换控制台显示/隐藏"""
        self.console_expanded = not self.console_expanded
        
        # 找到控制台切换按钮
        console_header = self.console_container.content.controls[0]
        toggle_btn = console_header.content.controls[3]
        
        if self.console_expanded:
            # 展开控制台
            self.console_container.content.controls[1].visible = True
            toggle_btn.icon = ft.Icons.KEYBOARD_ARROW_DOWN
            toggle_btn.tooltip = "最小化控制台"
            logger.info("控制台已展开")
        else:
            # 最小化控制台
            self.console_container.content.controls[1].visible = False
            toggle_btn.icon = ft.Icons.KEYBOARD_ARROW_UP
            toggle_btn.tooltip = "展开控制台"
            logger.info("控制台已最小化")
        
        self.page.update()

    def _toggle_auth(self, e):
        if not self.auth_status:  # 仅在未授权时触发授权流程
            start_authorize_flow()
            self.page.snack_bar = ft.SnackBar(ft.Text("已打开浏览器，请完成授权..."))
            self.page.snack_bar.open = True
            self.page.update()
            self.page.run_task(self._check_authorization)
        else:
            logger.info("已存在有效授权，无需重新授权")

    async def _check_authorization(self):
        from modes.feishu.localserver import code_queue
        logger.info("开始检测授权状态...")

        try:
            #await asyncio.wait_for(code_event.wait(), timeout=120)
            last_code = await asyncio.wait_for(code_queue.get(), timeout=120)
            logger.info(f"获取到的授权代码: {last_code}")  # 确认last_code值
            if not last_code:
                raise ValueError("未收到授权代码")
        except asyncio.TimeoutError:
            logger.warning("授权超时")
            self.page.snack_bar = ft.SnackBar(ft.Text("授权超时，请重试"))
            self.page.snack_bar.open = True
            self.page.update()
            return

        try:
            token, expires_in = exchange_code_for_token(last_code)
            expire_time = int(time.time()) + expires_in - 1000
            self.token_storage.set("user_token", token)
            self.token_storage.set("expire_time", expire_time)
            self.access_token = token
            logger.success(f"飞书授权成功! token={token[:10]}...，到期时间={expire_time}")
            self._set_auth_status(True)
            self.page.snack_bar = ft.SnackBar(ft.Text("飞书授权成功 ✅"))
            self.page.snack_bar.open = True
            self.page.update()

        except Exception as e:
            self.page.snack_bar = ft.SnackBar(ft.Text(f"授权失败: {e}"))
            self.page.snack_bar.open = True
            self.page.update()
            logger.error(f"授权失败: {e}")


    def _set_auth_status(self, success: bool):
        """更新授权状态显示"""
        icon_control = self.auth_indicator.content.controls[0]
        text_control = self.auth_indicator.content.controls[1]

        if success:
            icon_control.name = ft.Icons.LOCK_OPEN
            icon_control.color = ft.Colors.GREEN_400
            text_control.value = "已授权"
            text_control.color = ft.Colors.GREEN_400
            self.auth_indicator.border = ft.border.all(1, ft.Colors.GREEN_400)
            self.auth_status = True
        else:
            icon_control.name = ft.Icons.LOCK
            icon_control.color = ft.Colors.RED_400
            text_control.value = "未授权"
            text_control.color = ft.Colors.RED_400
            self.auth_indicator.border = ft.border.all(1, ft.Colors.RED_400)
            self.auth_status = False

        self.page.update()

    async def _start_local_server(self):
        import modes.feishu.localserver as localserver
        await localserver.run_server_async()

def main(page: ft.Page):
    """应用程序入口"""
    MainApp(page)


if __name__ == "__main__":
    ft.app(target=main)