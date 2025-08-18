from flask import Flask, render_template, request, jsonify,send_from_directory
import json
from apscheduler.schedulers.background import BackgroundScheduler
from datetime import datetime
from linebot import LineBotApi
from linebot.models import TextSendMessage
from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage, TextSendMessage, FileMessage
from dotenv import load_dotenv
import os
ENV = './.env' 
load_dotenv(dotenv_path=ENV)

app = Flask(__name__)

# ===== LINE 設定 =====
CHANNEL_ACCESS_TOKEN = os.getenv('CHANNEL_ACCESS_TOKEN')  # Messaging API Channel Access Token
CHANNEL_SECRET = os.getenv('CHANNEL_SECRET')
TEMP='/home/kingzaeip1/dailyreport/temp'
PNG='/home/kingzaeip1/dailyreport/static/img'
app.config['TEMP'] = TEMP
app.config['PNG'] = PNG
line_bot_api = LineBotApi(CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(CHANNEL_SECRET)
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
            per['LINE']=data['editLINE']
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
@app.route("/files/<path:filename>")
def serve_file(filename):
    return send_from_directory(app.config['TEMP'], filename, as_attachment=True)
@app.route("/png/<path:filename>")
def png_file(filename):
    return send_from_directory(app.config['PNG'], filename, as_attachment=True)
#================LINE WEBHOOK=====================
@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers['X-Line-Signature']
    body = request.get_data(as_text=True)

    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)

    return 'OK'

# 發送檔案下載連結
def send_excel_link(user_id, file_name):
    # 假設 Flask 跑在 localhost:5000
    print(file_name)
    file_url = f"https://cf23fc37feab.ngrok-free.app/files/{file_name}"
    file_message = FileMessage(
        original_content_url=file_url,  # 檔案實際 URL
        file_name=file_name,
        preview_image_url="https://cf23fc37feab.ngrok-free.app/png/logo.png",
        file_size = os.path.getsize(os.path.join(app.config['TEMP'], file_name))  # 這裡填檔案大小（bytes），可用 os.path.getsize() 自動取得
    )
   # line_bot_api.push_message(user_id, TextSendMessage(text=f"您的檔案下載連結：{file_url}"))
    line_bot_api.push_message(user_id, file_message)
# ====== 使用者加好友事件 (FollowEvent) ======
# @handler.add(FollowEvent)
# def handle_follow(event):
#     user_id = event.source.user_id
#     print("新加入的使用者 ID:", user_id)

#     # 可以回覆一則歡迎訊息
#     line_bot_api.reply_message(
#         event.reply_token,
#         TextSendMessage(text=f"歡迎加入！你的ID是 {user_id}")
#     )


# ====== 使用者傳訊息事件 ======
@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    user_id = event.source.user_id
    user_text = event.message.text

    print(f"收到來自 {user_id} 的訊息: {user_text}")
    if user_text[:2] == "工號":
        rest = user_text[2:]          # 取 "工號" 後面的字

        if rest[0].lower() == "a" or rest[0].lower() == "A":
            result = rest
            # 回覆訊息
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text=f"您的工號是: {result}\n你的ID是: {user_id}")
            )
    elif user_text=='Data':
        send_excel_link('Ue8115fd6e2a0ffb3170fa8a0949ce4b9','testdata.xlsx')
        




# 時間設定
@app.route("/set_time", methods=["POST"])
def set_time():
    hour = int(request.form.get("hour"))
    minute = int(request.form.get("minute"))
    settings["hour"] = hour
    settings["minute"] = minute
    return "更新成功", 200



# # ===== LINE 推送功能 =====
# def push_excel_link():
#     for uid, roles in permissions.items():
#         for role in roles:
#             file_name = f"{role}_report.xlsx"
#             file_url = f"http://your-public-ip:5000/files/{file_name}"  # 需要改成你 public IP 或 ngrok
#             try:
#                 line_bot_api.push_message(uid, TextSendMessage(text=f"您的檔案下載連結：{file_url}"))
#             except Exception as e:
#                 print(f"推送失敗 {uid}: {e}")

# # ===== APScheduler 排程 =====
# scheduler = BackgroundScheduler()
# scheduler.add_job(
#     push_excel_link,
#     'cron',
#     hour=settings["hour"],
#     minute=settings["minute"]
# )
# scheduler.start()

# ===== Flask 啟動 =====
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8018)
