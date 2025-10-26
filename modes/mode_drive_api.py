"""
mode_drive_api.py
飞书电子表格相关信息API获取
spreadsheetToken
sheetId
"""
import requests
from core.env import SPREADSHEET_TOKEN

def get_spreadsheetToken(access_token, page_size = 50):
    url = f"https://open.feishu.cn/open-apis/drive/v1/files?"
    f"page_size={page_size}&user_id_type=user_id"

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    """
        try:
        resp = requests.get(url, headers=headers)
        data = resp.json()
        code = data.get("code")
        message = data.get("msg")
        files = data.get("data").get("files")
        if code == 0:
            return files
        else:
            print(f"Error code{code}, {message}")
            raise Exception(f"Error code{code}, {message}")
    except Exception as e:
        print(e)
        raise e
    """
    return [
        {
            "name": f"表_{i + 1}",  # i 从 0 开始，所以 +1
            "parent_token": "fldcnCEG903UUB4fUqfysdabcef",
            "token": token,
            "type": "sheet"
        }
        for i, token in enumerate(SPREADSHEET_TOKEN)
    ]
"""
{
    "code":0,
    "data":{
        "files":[
            {
                "name":"test docx",
                "parent_token":"fldcnCEG903UUB4fUqfysdabcef",
                "token":"doxcntan34DX4QoKJu7jJyabcef",
                "type":"docx",
                "created_time":"1679295364",
                "modified_time":"1679295364",
                "owner_id":"ou_20b31734443364ec8a1df89fdf325b44",
                "url":"https://feishu.cn/docx/doxcntan34DX4QoKJu7jJyabcef"
            },
            {
                "name":"test sheet",
                "parent_token":"fldcnCEG903UUB4fUqfysdabcef",
                "token":"shtcnOko1Ad0HU48HH8KHuabcef",
                "type":"sheet",
                "created_time":"1679295364",
                "modified_time":"1679295364",
                "owner_id":"ou_20b31734443364ec8a1df89fdf325b44",
                "url":"https://feishu.cn/sheets/shtcnOko1Ad0HU48HH8KHuabcef"
            }
        ],
        "has_more":false
    },
    "msg":"success"
}
"""


def get_spreadsheet_Id(access_token, spreadsheet_token):
    url = f"https://open.feishu.cn/open-apis/sheets/v3/spreadsheets/{spreadsheet_token}/sheets/query"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    try:
        resp = requests.get(url, headers=headers)
        data = resp.json()
        code = data.get("code")
        message = data.get("msg")
        sheets = data.get("data").get("sheets")
        if code == 0:
            return sheets
        else:
            print(f"Error code{code}, {message}")
            raise Exception(f"Error code{code}, {message}")
    except Exception as e:
        raise e

"""
{
    "code": 0,
    "msg": "success",
    "data": {
        "sheets": [
            {
                "sheet_id": "sxj5ws",
                "title": "title",
                "index": 0,
                "hidden": false,
                "grid_properties": {
                    "frozen_row_count": 0,
                    "frozen_column_count": 0,
                    "row_count": 200,
                    "column_count": 20
                },
                "resource_type": "sheet",
                "merges": [
                    {
                        "start_row_index": 0,
                        "end_row_index": 0,
                        "start_column_index": 0,
                        "end_column_index": 0
                    }
                ]
            }
        ]
    }
}
"""