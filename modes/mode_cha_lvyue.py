import requests
import webbrowser
from core.env import SPREADSHEET_TOKEN, SHEET_ID
def get_table_filter(access_token, spreadsheet_token, sheet_id):
    url = f"https://open.feishu.cn/open-apis/sheets/v3/spreadsheets/{spreadsheet_token}/sheets/{sheet_id}/filter"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    try:
        resp = requests.get(url, headers=headers)
        data = resp.json()
        if data.get("code") == 0:
            return data.get("data").get("sheet_filter_info").get("range")
        else:
            raise Exception(data.get("msg"))
    except Exception as e:
        raise e

def get_table_value(access_token, spreadsheet_token, value_range):
    url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{spreadsheet_token}/values/{value_range}"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    try:
        resp = requests.get(url, headers=headers)
        data = resp.json()
        if data.get("code") == 0:
            return data.get("data").get("valueRange").get("values")
        else:
            raise Exception(data.get("msg"))
    except Exception as e:
        raise e


