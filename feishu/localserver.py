# localserver.py
from flask import Flask, request
import threading

app = Flask(__name__)
last_code = None

@app.route('/callback')
def callback():
    global last_code
    code = request.args.get('code')
    state = request.args.get('state')
    last_code = code
    print(f"[Callback] Got code: {code}, state: {state}")
    return "授权成功！您可以回到应用中继续操作。"

def run_server():
    app.run(host='0.0.0.0', port=3000, debug=False)
