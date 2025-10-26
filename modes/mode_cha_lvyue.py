import requests
import webbrowser
from core.env import SPREADSHEET_TOKEN, SHEET_ID
def get_table_filter(access_token, spreadsheet_token, sheet_id):
    url = f"https://open.feishu.cn/open-apis/sheets/v3/spreadsheets/{spreadsheet_token}/sheets/{sheet_id}"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    try:
        resp = requests.get(url, headers=headers)
        data = resp.json()
        if data.get("code") == 0:
            cols = data.get("data").get("sheet").get("grid_properties").get("column_count")
            rows = data.get("data").get("sheet").get("grid_properties").get("row_count")
            return f"{sheet_id}!{get_range_str(rows,cols)}"
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

def column_index_to_letter(idx: int) -> str:
    """
    将 1-based 列号转换为 Excel 列字母
    1 -> A, 26 -> Z, 27 -> AA, 28 -> AB ...
    """
    letters = ""
    while idx > 0:
        idx, remainder = divmod(idx - 1, 26)
        letters = chr(65 + remainder) + letters  # 65 -> 'A'
    return letters

def get_range_str(row_count: int, column_count: int) -> str:
    """
    根据行数和列数生成 Excel/Sheets 范围字符串
    """
    start_cell = "A1"
    end_column = column_index_to_letter(column_count)
    end_row = row_count
    return f"{start_cell}:{end_column}{end_row}"


