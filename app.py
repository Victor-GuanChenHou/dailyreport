from flask import Flask, render_template, request, jsonify,send_from_directory
import json
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from flask import Flask, request, abort
from linebot.v3 import (WebhookHandler)
from linebot.v3.exceptions import (InvalidSignatureError)
from linebot.v3.messaging import (Configuration, ApiClient,MessagingApi,ReplyMessageRequest,TextMessage)
from linebot.v3.webhooks import (MessageEvent,TextMessageContent)
from linebot.v3.messaging.models import (FlexBubble,FlexBox,FlexText,FlexMessage,PushMessageRequest,TemplateMessage,ButtonsTemplate,PostbackAction,MessageAction,URIAction)
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from dotenv import load_dotenv
import math
import os
ENV = './.env' 
load_dotenv(dotenv_path=ENV)

app = Flask(__name__)

# ===== LINE 設定 =====
CHANNEL_ACCESS_TOKEN = os.getenv('CHANNEL_ACCESS_TOKEN')  # Messaging API Channel Access Token
CHANNEL_SECRET = os.getenv('CHANNEL_SECRET')

TEMP='temp'
PNG='static/img'
FOLDER='static/file'
app.config['TEMP'] = TEMP
app.config['PNG'] = PNG
app.config['FOLDER'] = FOLDER
configuration = Configuration(access_token=CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(CHANNEL_SECRET)
api_client = ApiClient(configuration)
line_bot_api = MessagingApi(api_client)
global  last_setting
with open("settings.json", "r", encoding="utf-8") as f:
        last_settings = json.load(f)
last_setting=last_settings[0]
# ===== 全域資料 =====

settings = {"hour": 9, "minute": 0}  # 每日推送時間
def update_job():
    """檢查設定是否改變，更新排程"""
    global current_job, last_setting
    with open("settings.json", "r", encoding="utf-8") as f:
        settings = json.load(f)
    setting=settings[0]
    # 判斷是否需要更新 job
    if last_setting.get("hour") != setting.get("hour"):
        hour = setting.get("hour", 9)

        # 刪掉舊 job
        if current_job:
            scheduler.remove_job(current_job.id)

        # 建立新 job
        trigger = CronTrigger(hour=hour, minute=0)
        current_job = scheduler.add_job(send_message, trigger)
        print(f"[{datetime.now()}] 更新排程: 每天 {hour}:00 發送訊息")

    last_setting = setting
def send_message():
    """發送訊息任務"""
    with open("settings.json", "r", encoding="utf-8") as f:
        setting = json.load(f)
    message = setting.get("message", "預設訊息")
    print(f"[{datetime.now()}] 發送訊息: {message}")
def excelmake(user_id,day,data,start):#工號 日期資料 完整資料 資料excel期始位置
    with open("permissions.json", "r", encoding="utf-8") as f:
        permission = json.load(f)
    for per in permission:
        if per['user_id']==user_id:
            userdata=per
    user_folder = os.path.join(app.config['FOLDER'], user_id)
    if not os.path.exists(user_folder):
        os.makedirs(user_folder)
    end_date = datetime.strptime(day, "%Y-%m-%d")
    month = end_date.month
    dataday = end_date.day
    date_range_str = f"{month}/1 ~ {month}/{dataday}"
    ytd_range_str=f"1/1 ~ {month}/{dataday}"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"{day}Daily Report"

    # 日期 & 店數
    ws.merge_cells("A1:B1")
    ws["A1"] = day
    ws["A1"].alignment = Alignment(horizontal="center")
    ws["A2"] = "店數"
    ws["B2"] = len(data)-2
    # 標題顏色
    sales_fill = PatternFill("solid", fgColor="800000")   # 暗紅
    tc_fill = PatternFill("solid", fgColor="006666")      # 藍綠
    ta_fill = PatternFill("solid", fgColor="660066")      # 紫色
    header_font = Font(bold=True, color="FFFFFF")
    # 寫 Daily Sales 標題
    ws.merge_cells(f"C{start}:E{start}")
    ws[f"C{start}"] = "Daily Sales"
    ws[f"C{start}"].fill = sales_fill
    ws[f"C{start}"].font = header_font
    ws[f"C{start}"].alignment = Alignment(horizontal="center")
    ws[f"C{start+1}"], ws[f"D{start+1}"], ws[f"E{start+1}"] = "CY", "PY", "Index"
    for col in ["C", "D", "E"]:
        ws[f"{col}{start+1}"].fill = sales_fill
        ws[f"{col}{start+1}"].font = header_font
        ws[f"{col}{start+1}"].alignment = Alignment(horizontal="center")
    # 寫 Daily TC 標題
    ws.merge_cells(f"F{start}:H{start}")
    ws[f"F{start}"] = "Daily TC"
    ws[f"F{start}"].fill = tc_fill
    ws[f"F{start}"].font = header_font
    ws[f"F{start}"].alignment = Alignment(horizontal="center")

    ws[f"F{start+1}"], ws[f"G{start+1}"], ws[f"H{start+1}"] = "CY", "PY", "Index"
    for col in ["F", "G", "H"]:
        ws[f"{col}{start+1}"].fill = tc_fill
        ws[f"{col}{start+1}"].font = header_font
        ws[f"{col}{start+1}"].alignment = Alignment(horizontal="center")

    # 寫 Daily TA 標題
    ws.merge_cells(f"I{start}:K{start}")
    ws[f"I{start}"] = "Daily TA"
    ws[f"I{start}"].fill = ta_fill
    ws[f"I{start}"].font = header_font
    ws[f"I{start}"].alignment = Alignment(horizontal="center")

    ws[f"I{start+1}"], ws[f"J{start+1}"], ws[f"K{start+1}"] = "CY", "PY", "Index"
    for col in ["I", "J", "K"]:
        ws[f"{col}{start+1}"].fill = ta_fill
        ws[f"{col}{start+1}"].font = header_font
        ws[f"{col}{start+1}"].alignment = Alignment(horizontal="center")
    # 寫 MTD Sales 標題
    ws.merge_cells(f"L{start}:N{start}")
    ws[f"L{start}"] = f"MTD Sales({date_range_str})"
    ws[f"L{start}"].fill = sales_fill
    ws[f"L{start}"].font = header_font
    ws[f"L{start}"].alignment = Alignment(horizontal="center")
    ws[f"L{start+1}"], ws[f"M{start+1}"], ws[f"N{start+1}"] = "CY", "PY", "Index"
    for col in ["L", "M", "N"]:
        ws[f"{col}{start+1}"].fill = sales_fill
        ws[f"{col}{start+1}"].font = header_font
        ws[f"{col}{start+1}"].alignment = Alignment(horizontal="center")
    # 寫 MTD TC Sales 標題
    ws.merge_cells(f"O{start}:Q{start}")
    ws[f"O{start}"] = f"MTD TC Sales({date_range_str})"
    ws[f"O{start}"].fill = tc_fill
    ws[f"O{start}"].font = header_font
    ws[f"O{start}"].alignment = Alignment(horizontal="center")
    ws[f"O{start+1}"], ws[f"P{start+1}"], ws[f"Q{start+1}"] = "CY", "PY", "Index"
    for col in ["O", "P", "Q"]:
        ws[f"{col}{start+1}"].fill = tc_fill
        ws[f"{col}{start+1}"].font = header_font
        ws[f"{col}{start+1}"].alignment = Alignment(horizontal="center")
    # 寫 MTD TA Sales 標題
    ws.merge_cells(f"R{start}:T{start}")
    ws[f"R{start}"] = f"MTD TA Sales({date_range_str})"
    ws[f"R{start}"].fill = ta_fill
    ws[f"R{start}"].font = header_font
    ws[f"R{start}"].alignment = Alignment(horizontal="center")
    ws[f"R{start+1}"], ws[f"S{start+1}"], ws[f"T{start+1}"] = "CY", "PY", "Index"
    for col in ["R", "S", "T"]:
        ws[f"{col}{start+1}"].fill = ta_fill
        ws[f"{col}{start+1}"].font = header_font
        ws[f"{col}{start+1}"].alignment = Alignment(horizontal="center")
     # 寫 YTD Sales 標題
    ws.merge_cells(f"U{start}:W{start}")
    ws[f"U{start}"] = f"YTD Sales({ytd_range_str})"
    ws[f"U{start}"].fill = sales_fill
    ws[f"U{start}"].font = header_font
    ws[f"U{start}"].alignment = Alignment(horizontal="center")
    ws[f"U{start+1}"], ws[f"V{start+1}"], ws[f"W{start+1}"] = "CY", "PY", "Index"
    for col in ["U", "V", "W"]:
        ws[f"{col}{start+1}"].fill = sales_fill
        ws[f"{col}{start+1}"].font = header_font
        ws[f"{col}{start+1}"].alignment = Alignment(horizontal="center")
    # 寫 YTD TC Sales 標題
    ws.merge_cells(f"X{start}:Z{start}")
    ws[f"X{start}"] = f"YTD TC Sales({ytd_range_str})"
    ws[f"X{start}"].fill = tc_fill
    ws[f"X{start}"].font = header_font
    ws[f"X{start}"].alignment = Alignment(horizontal="center")
    ws[f"X{start+1}"], ws[f"P{start+1}"], ws[f"Q{start+1}"] = "CY", "PY", "Index"
    for col in ["X", "Y", "Z"]:
        ws[f"{col}{start+1}"].fill = tc_fill
        ws[f"{col}{start+1}"].font = header_font
        ws[f"{col}{start+1}"].alignment = Alignment(horizontal="center")
    # 寫 YTD TA Sales 標題
    ws.merge_cells(f"AA{start}:AC{start}")
    ws[f"AA{start}"] = f"YTD TA Sales({ytd_range_str})"
    ws[f"AA{start}"].fill = ta_fill
    ws[f"AA{start}"].font = header_font
    ws[f"AA{start}"].alignment = Alignment(horizontal="center")
    ws[f"AA{start+1}"], ws[f"AB{start+1}"], ws[f"AC{start+1}"] = "CY", "PY", "Index"
    for col in ["AA", "AB", "AC"]:
        ws[f"{col}{start+1}"].fill = ta_fill
        ws[f"{col}{start+1}"].font = header_font
        ws[f"{col}{start+1}"].alignment = Alignment(horizontal="center")
    row=start+2
    for r in data:
        ws[f"A{row}"] = r[0]
        ws[f"B{row}"] = r[1]
        ws[f"C{row}"] = r[2]
        ws[f"C{row}"].number_format = "#,##0"
        ws[f"D{row}"] = r[3]
        ws[f"D{row}"].number_format = "#,##0"
        ws[f"E{row}"] = r[4]
        ws[f"E{row}"].number_format = "#,##0"
        ws[f"F{row}"] = r[5]
        ws[f"F{row}"].number_format = "#,##0"
        ws[f"G{row}"] = r[6]
        ws[f"G{row}"].number_format = "#,##0"
        ws[f"H{row}"] = r[7]
        ws[f"H{row}"].number_format = "#,##0"
        ws[f"I{row}"] = r[8]
        ws[f"I{row}"].number_format = "#,##0"
        ws[f"J{row}"] = r[9]
        ws[f"J{row}"].number_format = "#,##0"
        ws[f"K{row}"] = r[10]
        ws[f"K{row}"].number_format = "#,##0"
        ws[f"L{row}"] = r[11]
        ws[f"L{row}"].number_format = "#,##0"
        ws[f"M{row}"] = r[12]
        ws[f"M{row}"].number_format = "#,##0"
        ws[f"N{row}"] = r[13]
        ws[f"N{row}"].number_format = "#,##0"
        ws[f"O{row}"] = r[14]
        ws[f"O{row}"].number_format = "#,##0"
        ws[f"P{row}"] = r[15]
        ws[f"P{row}"].number_format = "#,##0"
        ws[f"Q{row}"] = r[16]
        ws[f"Q{row}"].number_format = "#,##0"
        ws[f"R{row}"] = r[17]
        ws[f"R{row}"].number_format = "#,##0"
        ws[f"S{row}"] = r[18]
        ws[f"S{row}"].number_format = "#,##0"
        ws[f"T{row}"] = r[19]
        ws[f"T{row}"].number_format = "#,##0"
        ws[f"U{row}"] = r[20]
        ws[f"U{row}"].number_format = "#,##0"
        ws[f"V{row}"] = r[21]
        ws[f"V{row}"].number_format = "#,##0"
        ws[f"W{row}"] = r[22]
        ws[f"W{row}"].number_format = "#,##0"
        ws[f"X{row}"] = r[23]
        ws[f"X{row}"].number_format = "#,##0"
        ws[f"Y{row}"] = r[24]
        ws[f"Y{row}"].number_format = "#,##0"
        ws[f"Z{row}"] = r[25]
        ws[f"Z{row}"].number_format = "#,##0"
        ws[f"AA{row}"] = r[26]
        ws[f"AA{row}"].number_format = "#,##0"
        ws[f"AB{row}"] = r[27]
        ws[f"AB{row}"].number_format = "#,##0"
        ws[f"AC{row}"] = r[28]
        ws[f"AC{row}"].number_format = "#,##0"
        row += 1

    # 美化欄寬
    for col in range(1, 30):
        ws.column_dimensions[get_column_letter(col)].width = 15
    wb.save(f"{user_folder}/{day}daily_report.xlsx")
def Send_EMAIL(user_id,day):
    # 郵件內容設定
    sender_email = os.getenv('MAIL')
    password = os.getenv('MAIL_PW')
    with open("permissions.json", "r", encoding="utf-8") as f:
        permission = json.load(f)
    for per in permission:
        if per['user_id']==user_id:
            email=per['email']
    receiver_email=email
    subject = f"{day}日報表"
    filepath = os.path.join(app.config['FOLDER'], user_id, f"{day}daily_report.xlsx")


    # 建立郵件物件
    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = subject
    # 郵件主體
    body_html = f"""
    <html>
    <head><meta charset="utf-8"></head>
    <body>
    <p>附件為{day}日報表再請參考</p>
    </body>
    </html>
    """
    message.attach(MIMEText(body_html, "html"))

    # 加入 Excel 附件
    if os.path.exists(filepath):
        with open(filepath, "rb") as f:
            part = MIMEApplication(f.read(), Name=os.path.basename(filepath))
        part['Content-Disposition'] = f'attachment; filename="{os.path.basename(filepath)}"'
        message.attach(part)
    else:
        print(f"警告：檔案不存在 -> {filepath}")



    try:
        # 建立與 Gmail SMTP 伺服器的連線 (使用 SSL)
        with smtplib.SMTP_SSL("mail.kingza.com.tw", 465) as server:
            if not (isinstance(email, float) and math.isnan(email)):
                server.login(sender_email, password)
                server.sendmail(sender_email, receiver_email, message.as_string())
                print("郵件寄送成功！")

    except Exception as e:
        print(f"發生錯誤：{e}")
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
@app.route("/home")
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
        "name": data.get("name"),
        "dept": data.get("dept")
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
            per['dept']=data['dept']
            break
    with open("store.json", "w", encoding="utf-8") as f:
        json.dump(store, f, ensure_ascii=False, indent=4)
    return jsonify({"success": True, "message": "資料已儲存"})
@app.route("/setting")
def setting():
    with open("settings.json", "r", encoding="utf-8") as f:
        setting = json.load(f)
    return render_template("setting.html",setting=setting,settings=settings)
@app.route("/editsetting", methods=['POST'])
def editsetting():
    with open("settings.json", "r", encoding="utf-8") as f:
        setting = json.load(f)
    data = request.get_json()
    for per in setting:
        per['hour']=data['hour']
        per['minute']=data['minute']
        per['ngrokid']=data['ngrokid']
        break
    with open("settings.json", "w", encoding="utf-8") as f:
        json.dump(setting, f, ensure_ascii=False, indent=4)
    return jsonify({"success": True, "message": "資料已儲存"})
@app.route("/files/<user_id>/<path:filename>")
def serve_file(user_id,filename):
    folder_path = os.path.join(app.config['FOLDER'], user_id)
    return send_from_directory(folder_path, filename, as_attachment=True)
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
def send_table(user_id):
    top_departments = [
        {"name": "業務", "count": 120},
        {"name": "工程", "count": 150},
        {"name": "行銷", "count": 80},
        {"name": "客服", "count": 30},
        {"name": "財務", "count": 60},
        {"name": "設計", "count": 45},
        {"name": "採購", "count": 25},
        {"name": "物流", "count": 35},
        {"name": "法務", "count": 20},
        {"name": "行政", "count": 40}
    ]

    # 生成 Flex Message 內容
    flex_contents = {
        "type": "bubble",
        "body": {
            "type": "box",
            "layout": "vertical",
            "contents": [
                {"type": "text", "text": "📊 每日報表 Top 10", "weight": "bold", "size": "xl"},
                {
                    "type": "box",
                    "layout": "vertical",
                    "contents": [
                        {"type": "text", "text": f"{d['name']}       {d['count']}"} 
                        for d in top_departments
                    ]
                }
            ]
        }
    }
    
     # 產生 FlexText 內容
    text_items = [FlexText(text=f"{d['name']}       {d['count']}") for d in top_departments]

    # Box 包起來
    body_box = FlexBox(layout="vertical", contents=[FlexText(text="📊 每日報表 Top 10", weight="bold", size="xl")] + text_items)

    # Bubble
    bubble = FlexBubble(body=body_box)

    # FlexMessage
    flex_message = FlexMessage(alt_text="每日報表", contents=bubble)

    # 推播
    line_bot_api.push_messages(
        to=user_id,
        messages=[flex_message]
    )
def send_excel_button(user_id, file_name,day):
    with open("settings.json", "r", encoding="utf-8") as f:
        setting = json.load(f)
    setting = setting[0]
    
    with open("permissions.json", "r", encoding="utf-8") as f:
        permission = json.load(f)
    for per in permission:
        if per['LINE']==user_id:
            ID=per['user_id']
    file_url = f"https://{setting['ngrokid']}/files/{ID}/{file_name}"
    buttons_template = ButtonsTemplate(
        thumbnail_image_url=f"https://{setting['ngrokid']}/png/logo.png",
        title="日報表",
        text=day,
        actions=[
            URIAction(label="Download", uri=file_url),
        ]
    )
    message = TemplateMessage(
        alt_text="ButtonsTemplate",
        template=buttons_template
    )
    line_bot_api.push_message(
        PushMessageRequest(
            to=user_id,
            messages=[message]
        )
    )
# ====== 使用者傳訊息事件 ======
@handler.add(MessageEvent, message=TextMessageContent)
def handle_message(event):
    user_id = event.source.user_id
    user_text = event.message.text

    print(f"收到來自 {user_id} 的訊息: {user_text}")
    if user_text[:2] == "工號":
        rest = user_text[2:]          # 取 "工號" 後面的字

        if rest[0].lower() == "a" or rest[0].lower() == "A":
            result = rest
            # 回覆訊息
            line_bot_api.reply_message_with_http_info(
                ReplyMessageRequest(
                    reply_token=event.reply_token,
                    messages=[TextMessage(text=f"您的工號是: {result}\n你的ID是: {user_id}")]
                    )
            )
    elif user_text=='Data':
        data = [
            ["全品牌", 19094808, "", "", 28896, "", "", 661, "", "", "", 28896, "", "", 661, "", "", 28896, "", "", 661, "", "", 28896, "", "", 661, "", 661],
            ["Total", 19094808, "", "", 28896, "", "", 661, "", "", "", 28896, "", "", 661, "", "", 28896, "", "", 661, "", "", 28896, "", "", 661, "", 661],
            ["蘭城新月", 19094808, "", "", 28896, "", "", 661, "", "", "", 28896, "", "", 661, "", "", 28896, "", "", 661, "", "", 28896, "", "", 661, "", 661],
            ["信義威秀", 19094808, "", "", 28896, "", "", 661, "", "", "", 28896, "", "", 661, "", "", 28896, "", "", 661, "", "", 28896, "", "", 661, "", 661],
            ["廣三SOGO", 19094808, "", "", 28896, "", "", 661, "", "", "", 28896, "", "", 661, "", "", 28896, "", "", 661, "", "", 28896, "", "", 661, "", 661],
            ["板橋環球", 19094808, "", "", 28896, "", "", 661, "", "", "", 28896, "", "", 661, "", "", 28896, "", "", 661, "", "", 28896, "", "", 661, "", 661],
            ["高雄義大", 19094808, "", "", 28896, "", "", 661, "", "", "", 28896, "", "", 661, "", "", 28896, "", "", 661, "", "", 28896, "", "", 661, "", 661],
            ["左營環球", 19094808, "", "", 28896, "", "", 661, "", "", "", 28896, "", "", 661, "", "", 28896, "", "", 661, "", "", 28896, "", "", 661, "", 661],
        ]

        excelmake("A14176",'2025-08-10',data,5)
        date='2025-08-10'
        send_excel_button('Ue8115fd6e2a0ffb3170fa8a0949ce4b9',f'{date}daily_report.xlsx')
    elif user_text=='Data2':
        send_table('Ue8115fd6e2a0ffb3170fa8a0949ce4b9')

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
scheduler = BackgroundScheduler()
current_job = None
scheduler.add_job(update_job, 'interval', minutes=1)
scheduler.start()
if __name__ == "__main__":

    app.run(host="0.0.0.0", port=8018)
