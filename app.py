import os
import imaplib
import email
from email.header import decode_header
from flask import Flask, render_template, request, redirect, url_for, session
from dotenv import load_dotenv
import math

# ↓↓↓ タイムゾーンと言語処理のためのライブラリをインポート ↓↓↓
from datetime import datetime
from zoneinfo import ZoneInfo
from email.utils import parsedate_to_datetime
# ↑↑↑ ここまで ↑↑↑

load_dotenv()

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY")

IMAP_SERVERS = {
    "gmail.com": "imap.gmail.com",
    "outlook.com": "outlook.office365.com",
    "hotmail.com": "outlook.office365.com",
    "live.com": "outlook.office365.com",
    "yahoo.co.jp": "imap.mail.yahoo.co.jp",
    "icloud.com": "imap.mail.me.com",
    "me.com": "imap.mail.me.com",
}

def decode_str(s, encoding='utf-8'):
    if isinstance(s, bytes):
        return s.decode(encoding if encoding else 'utf-8', errors='ignore')
    return s

# ↓↓↓ 日時を日本時間に変換する新しい関数を追加 ↓↓↓
def format_date_to_jst(date_string):
    """メールの日時文字列を解析し、日本のタイムゾーンに変換して整形する関数"""
    if not date_string:
        return "日時不明"
    try:
        # メールヘッダーの日時文字列をdatetimeオブジェクトに変換
        dt_object = parsedate_to_datetime(date_string)
        
        # タイムゾーンが未設定の場合（稀）、UTCとして扱う
        if dt_object.tzinfo is None:
            dt_object = dt_object.replace(tzinfo=ZoneInfo("UTC"))

        # 日本標準時（JST）のタイムゾーンを定義
        jst = ZoneInfo("Asia/Tokyo")
        
        # JSTに変換
        dt_jst = dt_object.astimezone(jst)
        
        # 見やすい形式の文字列に変換して返す (例: 2025-10-07 09:35)
        return dt_jst.strftime("%Y-%m-%d %H:%M")
    except Exception as e:
        print(f"日付の解析に失敗: {e}, 元の日付: {date_string}")
        # 解析に失敗した場合は、元の文字列をそのまま返す
        return date_string
# ↑↑↑ ここまで ↑↑↑

def get_email_body(msg):
    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            content_disposition = str(part.get("Content-Disposition"))
            if "attachment" not in content_disposition:
                if content_type == "text/html":
                    return decode_str(part.get_payload(decode=True), part.get_content_charset())
                elif content_type == "text/plain" and not body:
                    body = decode_str(part.get_payload(decode=True), part.get_content_charset())
    else:
        body = decode_str(msg.get_payload(decode=True), msg.get_content_charset())
    return body

@app.route("/")
def home():
    if "credentials" in session:
        return redirect(url_for("inbox"))
    return redirect(url_for("login"))

@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        email_address = request.form["email"]
        password = request.form["password"]
        try:
            domain = email_address.split('@')[1]
        except IndexError:
            domain = None
        imap_server = IMAP_SERVERS.get(domain)

        if not imap_server:
            error = f"対応していない、または不明なメールサービスです: ({domain})"
        else:
            try:
                mail_test = imaplib.IMAP4_SSL(imap_server)
                mail_test.login(email_address, password)
                mail_test.logout()
                session["credentials"] = {
                    "imap_server": imap_server,
                    "email_address": email_address,
                    "password": password
                }
                return redirect(url_for("inbox"))
            except Exception:
                error = "ログインに失敗しました。メールアドレスやパスワード（またはアプリパスワード）を確認してください。"
    
    return render_template("login.html", error=error)

@app.route("/inbox")
def inbox():
    if "credentials" not in session:
        return redirect(url_for("login"))
    
    page = request.args.get('page', 1, type=int)
    EMAILS_PER_PAGE = 25
    
    creds = session["credentials"]
    try:
        mail = imaplib.IMAP4_SSL(creds["imap_server"])
        mail.login(creds["email_address"], creds["password"])
        mail.select("inbox")
        status, messages = mail.search(None, "ALL")
        email_ids = messages[0].split()
        
        total_emails = len(email_ids)
        total_pages = math.ceil(total_emails / EMAILS_PER_PAGE)
        start = (page - 1) * EMAILS_PER_PAGE
        end = start + EMAILS_PER_PAGE

        target_ids = list(reversed(email_ids))[start:end]

        emails = []
        for email_id_bytes in target_ids:
            status, msg_data = mail.fetch(email_id_bytes, "(BODY.PEEK[HEADER.FIELDS (SUBJECT FROM DATE)])")
            header_data = msg_data[0][1]
            msg = email.message_from_bytes(header_data)
            
            subject_header = decode_header(msg["Subject"])[0]
            from_header = decode_header(msg["From"])[0]

            emails.append({
                "id": email_id_bytes.decode(),
                "subject": decode_str(subject_header[0], subject_header[1]),
                "from": decode_str(from_header[0], from_header[1]),
                # ↓↓↓ この行を変更: format_date_to_jst関数を呼び出す ↓↓↓
                "date": format_date_to_jst(msg.get("Date"))
            })
        mail.logout()
        
        return render_template("inbox.html", emails=emails, email_address=creds["email_address"],
                               page=page, total_pages=total_pages)
    except Exception as e:
        print(f"受信トレイ表示エラー: {e}")
        session.pop("credentials", None)
        return redirect(url_for("login"))

@app.route("/view/<email_id>")
def view_email(email_id):
    if "credentials" not in session:
        return redirect(url_for("login"))
    
    creds = session["credentials"]
    try:
        mail = imaplib.IMAP4_SSL(creds["imap_server"])
        mail.login(creds["email_address"], creds["password"])
        mail.select("inbox")
        
        status, msg_data = mail.fetch(email_id, "(RFC822)")
        msg = email.message_from_bytes(msg_data[0][1])

        subject_header = decode_header(msg["Subject"])[0]
        from_header = decode_header(msg.get("From"))[0]

        email_details = {
            "subject": decode_str(subject_header[0], subject_header[1]),
            "from": decode_str(from_header[0], from_header[1]),
            # ↓↓↓ この行を変更: format_date_to_jst関数を呼び出す ↓↓↓
            "date": format_date_to_jst(msg.get("Date")),
            "body": get_email_body(msg)
        }
        
        mail.logout()
        return render_template("view.html", email=email_details)
    except Exception as e:
        print(f"メール表示エラー: {e}")
        return redirect(url_for("inbox"))

@app.route("/logout")
def logout():
    session.pop("credentials", None)
    return redirect(url_for("login"))

if __name__ == "__main__":
    app.run(debug=True)