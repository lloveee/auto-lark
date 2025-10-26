import requests
import webbrowser
from core.env import SPREADSHEET_TOKEN, SHEET_ID
def get_table_filter(access_token, filter_range, filter_col, filter_type, compare_type, expected):
    url = f"https://open.feishu.cn/open-apis/sheets/v3/spreadsheets/{SPREADSHEET_TOKEN}/sheets/{SHEET_ID}/filter"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    payload = {
        "range": f"{SHEET_ID}!{filter_range}",
        "col": filter_col,
        "condition": {
            "filter_type": filter_type,
            "compare_type": compare_type,
            "expected": expected
        }
    }
    try:
        resp = requests.post(url, headers=headers, json=payload)
        data = resp.json()
        if data.get("code") == 0:
            return data.get("data")
        else:
            raise Exception(data.get("msg"))
    except Exception as e:
        raise e


