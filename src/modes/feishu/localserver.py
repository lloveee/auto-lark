# localserver.py
from quart import Quart, request
import threading
import asyncio

app = Quart(__name__)
last_code = None
#code_event = asyncio.Event()
code_queue = asyncio.Queue()

@app.route('/callback')
async def callback():
    global last_code
    from asyncio import get_event_loop
    code = request.args.get('code')
    state = request.args.get('state')
    last_code = code
    print(f"[Callback] Got code: {code}, state: {state}")
    await code_queue.put(code)

    return "授权成功！您可以回到应用中继续操作。"

async def run_server_async():
    """直接异步运行 Quart 服务器"""
    # 绑定 0.0.0.0:3000
    # use_reloader=False 避免启动两次
    await app.run_task(host="0.0.0.0", port=3000)
