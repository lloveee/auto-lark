# feishu_auth.py
import requests
import webbrowser
import time
from modes.feishu.localserver import last_code
from core.env import APP_ID, APP_SECRET, REDIRECT_URI

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
            print(data)

            return data["access_token"], data["expires_in"]
        else:
            raise Exception(f"Feishu error: {data}")
    else:
        raise Exception(f"HTTP {resp.status_code}: {resp.text}")

def start_authorize_flow():
    url = get_authorize_url()
    webbrowser.open(url)