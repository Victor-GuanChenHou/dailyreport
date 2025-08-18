from flask import Flask, render_template, request, jsonify
import json
from apscheduler.schedulers.background import BackgroundScheduler
from datetime import datetime
from linebot import LineBotApi
from linebot.models import TextSendMessage
from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage, TextSendMessage, FollowEvent

app = Flask(__name__)

# ===== LINE 設定 =====
LINE_CHANNEL_ACCESS_TOKEN = "YOUR_CHANNEL_ACCESS_TOKEN"
line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)

# ===== 全域資料 =====

settings = {"hour": 9, "minute": 0}  # 每日推送時間


@app.route("/data")
def index():
    with open("permissions.json", "r", encoding="utf-8") as f:
        permissions = json.load(f)
    with open("store.json", "r", encoding="utf-8") as f:
        stores = json.load(f)
    return render_template("index.html", permissions=permissions,stores=stores,settings=settings)
@app.route("/adduser", methods=['POST'])
def adduser():
    with open("permissions.json", "r", encoding="utf-8") as f:
        permissions = json.load(f)
    data = request.get_json()
    permissions.append({
        "departments": data.get("storeValues", []),
        "email": data.get("email"),
        "user_id": data.get("user"),
        "name": data.get("name")
    })
    with open("permissions.json", "w", encoding="utf-8") as f:
        json.dump(permissions, f, ensure_ascii=False, indent=4)
    return jsonify({"success": True, "message": "資料已儲存"})
@app.route("/edituser", methods=['POST'])
def edituser():
    with open("permissions.json", "r", encoding="utf-8") as f:
        permissions = json.load(f)
    data = request.get_json()
    for per in permissions:
        if per['user_id']==data['editUser']:
            per['name']=data['editName']
            per['departments']=data['editStore']
            per['email']=data['editEmail']
            break
    with open("permissions.json", "w", encoding="utf-8") as f:
        json.dump(permissions, f, ensure_ascii=False, indent=4)
    return jsonify({"success": True, "message": "資料已儲存"})
@app.route("/deletuser", methods=['POST'])
def deletuser():
    with open("permissions.json", "r", encoding="utf-8") as f:
        permissions = json.load(f)
    data = request.get_json()
    newdata=[]
    for per in permissions:
        if per['user_id']==data['userid']:
            pass
        else:
            newdata.append(per)
    with open("permissions.json", "w", encoding="utf-8") as f:
        json.dump(newdata, f, ensure_ascii=False, indent=4)
    return jsonify({"success": True, "message": "資料已儲存"})
@app.route("/")
def home():
    return render_template("home.html" )
@app.route("/store")
def store():
    with open("store.json", "r", encoding="utf-8") as f:
        stores = json.load(f)
    return render_template("store.html",stores=stores,settings=settings)
@app.route("/addstore", methods=['POST'])
def addstore():
    with open("store.json", "r", encoding="utf-8") as f:
        store = json.load(f)
    data = request.get_json()
    store.append({
        "value": data.get("value"),
        "name": data.get("name")
    })
    with open("store.json", "w", encoding="utf-8") as f:
        json.dump(store, f, ensure_ascii=False, indent=4)
    return jsonify({"success": True, "message": "資料已儲存"})
@app.route("/deletstore", methods=['POST'])
def deletstore():
    with open("store.json", "r", encoding="utf-8") as f:
        store = json.load(f)
    data = request.get_json()
    newdata=[]
    for per in store:
        if per['value']==data['value']:
            pass
        else:
            newdata.append(per)
    with open("store.json", "w", encoding="utf-8") as f:
        json.dump(newdata, f, ensure_ascii=False, indent=4)
    return jsonify({"success": True, "message": "資料已儲存"})
@app.route("/editstore", methods=['POST'])
def editstore():
    with open("store.json", "r", encoding="utf-8") as f:
        store = json.load(f)
    data = request.get_json()
    for per in store:
        if per['value']==data['value']:
            per['name']=data['name']
            break
    with open("store.json", "w", encoding="utf-8") as f:
        json.dump(store, f, ensure_ascii=False, indent=4)
    return jsonify({"success": True, "message": "資料已儲存"})

# 時間設定
@app.route("/set_time", methods=["POST"])
def set_time():
    hour = int(request.form.get("hour"))
    minute = int(request.form.get("minute"))
    settings["hour"] = hour
    settings["minute"] = minute
    return "更新成功", 200



# ===== LINE 推送功能 =====
def push_excel_link():
    for uid, roles in permissions.items():
        for role in roles:
            file_name = f"{role}_report.xlsx"
            file_url = f"http://your-public-ip:5000/files/{file_name}"  # 需要改成你 public IP 或 ngrok
            try:
                line_bot_api.push_message(uid, TextSendMessage(text=f"您的檔案下載連結：{file_url}"))
            except Exception as e:
                print(f"推送失敗 {uid}: {e}")

# ===== APScheduler 排程 =====
scheduler = BackgroundScheduler()
scheduler.add_job(
    push_excel_link,
    'cron',
    hour=settings["hour"],
    minute=settings["minute"]
)
scheduler.start()

# ===== Flask 啟動 =====
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
