# feishu_sheet.py
import requests


def col_num_to_letter(n):
    """将列号转换为Excel列字母（1→A, 27→AA）"""
    result = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result


def append_to_sheet(access_token, spreadsheet_token, range_ref, values):
    """
    :param access_token: user_access_token
    :param spreadsheet_token: 表格 token
    :param sheet_id: 目标 sheet id
    :param values: 二维数组数据 [[A1, B1], [A2, B2]]
    """
    url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{spreadsheet_token}/values_append"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    payload = {
        "valueRange": {
            "range": range_ref,
            "values": values
        }
    }

    resp = requests.post(url, headers=headers, json=payload)
    data = resp.json()
    if data.get("code") == 0:
        print("✅ 数据追加成功！")
        return True
    else:
        print("❌ 追加失败：", data)
        print("响应内容:", data)
        return False
