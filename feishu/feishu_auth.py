# feishu_auth.py
import requests
import threading
import webbrowser
import time
from localserver import run_server, last_code

APP_ID = "cli_a871ad50e192100b"         # 替换成你的 app_id
APP_SECRET = "pLAlJXXXdtHYNnrVN0rwsbMgy1TD6OLO"        # 替换成你的 app_secret
REDIRECT_URI = "http%3A%2F%2Flocalhost%3A3000%2Fcallback"
def start_local_server():
    thread = threading.Thread(target=run_server, daemon=True)
    thread.start()
    time.sleep(1)  # 等待服务器启动

def get_authorize_url(state="STATE"):
    return (
        f"https://accounts.feishu.cn/open-apis/authen/v1/authorize"
        f"?client_id={APP_ID}&response_type=code&redirect_uri={REDIRECT_URI}&scope=docs:doc%20drive:drive%20sheets:spreadsheet&state={state}"
    )

def wait_for_code(timeout=120):
    print("等待用户在浏览器中授权...")
    for i in range(timeout):
        if last_code:
            return last_code
        time.sleep(1)
    raise TimeoutError("授权超时")

def exchange_code_for_token(code):
    url = "https://open.feishu.cn/open-apis/authen/v2/oauth/token"
    payload = {
        "grant_type": "authorization_code",
        "client_id": APP_ID,
        "client_secret": APP_SECRET,
        "code": code,
        "redirect_uri": REDIRECT_URI
    }
    resp = requests.post(url, json=payload)
    if resp.status_code == 200:
        data = resp.json()
        if data.get("code") == 0:
            print("成功获取 user_access_token！")

            return data["access_token"]
        else:
            raise Exception(f"Feishu error: {data}")
    else:
        raise Exception(f"HTTP {resp.status_code}: {resp.text}")

def start_authorize_flow():
    """仅启动本地服务 + 打开浏览器，不阻塞"""
    start_local_server()
    url = get_authorize_url()
    webbrowser.open(url)