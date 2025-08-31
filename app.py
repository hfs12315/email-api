from flask import Flask, request, jsonify
import requests
import imaplib
import email
from bs4 import BeautifulSoup
from email.header import decode_header
import chardet
import os
import traceback

# Zeabur环境变量
CLIENT_ID = os.environ.get('CLIENT_ID', '9e5f94bc-e8a4-4e73-b8be-63364c29d753')
TENANT_ID = os.environ.get('TENANT_ID', 'common')

app = Flask(__name__)

def get_new_access_token(refresh_token):
    """使用刷新令牌获取新的访问令牌"""
    try:
        refresh_token_data = {
            'grant_type': 'refresh_token',
            'refresh_token': refresh_token,
            'client_id': CLIENT_ID,
        }
        token_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
        response = requests.post(token_url, data=refresh_token_data, timeout=30)
        response.raise_for_status()
        return response.json().get("access_token")
    except Exception as e:
        print(f"Token refresh error: {str(e)}")
        return None

def generate_auth_string(user, token):
    """生成 OAuth2 授权字符串"""
    return f"user={user}\1auth=Bearer {token}\1\1"

def decode_mime_words(s):
    """解码邮件标题"""
    try:
        if not s:
            return ""
        decoded_fragments = decode_header(s)
        return ''.join([str(t[0], t[1] or 'utf-8') if isinstance(t[0], bytes) else str(t[0]) for t in decoded_fragments])
    except Exception as e:
        return str(s) if s else ""

def strip_html(content):
    """去除 HTML 标签"""
    try:
        soup = BeautifulSoup(content, "html.parser")
        return soup.get_text()
    except Exception as e:
        return content

def safe_decode(byte_content):
    """自动检测并解码字节数据"""
    try:
        if not byte_content:
            return ""
        result = chardet.detect(byte_content)
        encoding = result.get('encoding', 'utf-8')
        if encoding:
            return byte_content.decode(encoding)
        else:
            return byte_content.decode('utf-8', errors='ignore')
    except Exception as e:
        return str(byte_content)

def remove_extra_blank_lines(text):
    """去除多余空行"""
    try:
        lines = text.splitlines()
        return "\n".join(filter(lambda line: line.strip(), lines))
    except Exception as e:
        return text

def get_emails_from_folder(mail, folder_name, max_emails=5):
    """从指定文件夹获取邮件内容"""
    emails_data = []
    try:
        status, messages = mail.select(folder_name, readonly=True)
        if status != "OK":
            return emails_data

        status, message_ids = mail.search(None, 'ALL')
        if status != "OK":
            return emails_data

        message_list = message_ids[0].split()
        limited_messages = message_list[-max_emails:] if len(message_list) > max_emails else message_list

        for i, message_id in enumerate(limited_messages):
            try:
                status, msg_data = mail.fetch(message_id, '(RFC822)')
                if status != "OK":
                    continue

                for response_part in msg_data:
                    if isinstance(response_part, tuple):
                        msg = email.message_from_bytes(response_part[1])
                        subject = decode_mime_words(msg.get("subject", ""))
                        date = msg.get("date", "")
                        sender = decode_mime_words(msg.get("from", ""))
                        body = ""

                        if msg.is_multipart():
                            for part in msg.walk():
                                content_type = part.get_content_type()
                                content_disposition = str(part.get("Content-Disposition", ""))

                                if "attachment" not in content_disposition:
                                    if content_type == "text/plain":
                                        payload = part.get_payload(decode=True)
                                        if payload:
                                            body += safe_decode(payload)
                                    elif content_type == "text/html":
                                        payload = part.get_payload(decode=True)
                                        if payload:
                                            html_content = safe_decode(payload)
                                            body += strip_html(html_content)
                        else:
                            payload = msg.get_payload(decode=True)
                            if payload:
                                if msg.get_content_type() == "text/plain":
                                    body = safe_decode(payload)
                                elif msg.get_content_type() == "text/html":
                                    html_content = safe_decode(payload)
                                    body = strip_html(html_content)

                        body = remove_extra_blank_lines(body)
                        emails_data.append(
                            f"\n--- 文件夹: {folder_name} ---\n" \
                            f"邮件编号: {i + 1}\n" \
                            f"邮件主题: {subject}\n" \
                            f"收件时间: {date}\n" \
                            f"发件人: {sender}\n" \
                            f"邮件正文:\n{body[:400]}{'...' if len(body) > 400 else ''}\n" \
                            f"\n--------------------------------------------------\n"
                        )
            except Exception as e:
                continue

    except Exception as e:
        pass
    
    return emails_data

@app.route('/')
def index():
    """根路径，提供API说明"""
    return """
    <html>
    <head>
        <title>邮件获取API - Zeabur部署</title>
        <meta charset="UTF-8">
        <style>
            body { font-family: Arial, sans-serif; max-width: 800px; margin: 0 auto; padding: 20px; background: #f5f5f5; }
            .container { background: white; padding: 30px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
            h1 { color: #2c3e50; text-align: center; }
            .status { background: #d4edda; border: 1px solid #c3e6cb; padding: 15px; border-radius: 5px; margin: 20px 0; }
            .api-info { background: #e3f2fd; border: 1px solid #bbdefb; padding: 20px; border-radius: 5px; margin: 20px 0; }
            code { background: #f8f9fa; padding: 5px 10px; border-radius: 3px; font-family: monospace; }
        </style>
    </head>
    <body>
        <div class="container">
            <h1>📧 邮件获取API</h1>
            <div class="status">
                <h2>🎉 部署成功！</h2>
                <p>服务运行正常，亚洲地区访问优化</p>
                <p><strong>部署平台：</strong> Zeabur (亚洲专用)</p>
            </div>
            <div class="api-info">
                <h3>📡 API使用方法</h3>
                <p><strong>接口地址：</strong></p>
                <p><code>GET /get_emails?email_address=你的邮箱&refresh_token=你的令牌</code></p>
                <p><strong>健康检查：</strong> <a href="/health">/health</a></p>
            </div>
        </div>
    </body>
    </html>
    """

@app.route('/health')
def health():
    """健康检查接口"""
    return jsonify({
        "status": "ok", 
        "message": "服务运行正常", 
        "platform": "Zeabur",
        "region": "Asia",
        "version": "1.0"
    })

@app.route('/get_emails')
def get_emails_api():
    try:
        email_address = request.args.get('email_address')
        refresh_token = request.args.get('refresh_token')

        print(f"📧 收到请求 - 邮箱: {email_address}")

        if not email_address or not refresh_token:
            return jsonify({"error": "❌ 缺少必要参数: email_address 或 refresh_token"}), 400

        # 获取访问令牌
        print("🔑 正在获取访问令牌...")
        access_token = get_new_access_token(refresh_token)
        
        if not access_token:
            return jsonify({"error": "❌ 无法获取访问令牌，请检查refresh_token是否有效"}), 500

        print("✅ 访问令牌获取成功，正在连接邮箱...")

        # 连接IMAP服务器
        mail = imaplib.IMAP4_SSL('outlook.office365.com')
        auth_string = generate_auth_string(email_address, access_token)
        mail.authenticate('XOAUTH2', lambda x: auth_string)
        print("✅ 邮箱认证成功")

        # 获取邮件
        all_emails_text = []
        
        # 获取收件箱邮件
        print("📥 正在获取收件箱邮件...")
        inbox_emails = get_emails_from_folder(mail, "INBOX", 5)
        all_emails_text.extend(inbox_emails)
        print(f"✅ 获取到 {len(inbox_emails)} 封收件箱邮件")

        # 获取垃圾邮件
        print("🗑️ 正在获取垃圾邮件...")
        junk_emails = get_emails_from_folder(mail, "Junk", 3)
        if not junk_emails:
            junk_emails = get_emails_from_folder(mail, "Junk Email", 3)
        all_emails_text.extend(junk_emails)
        print(f"✅ 获取到 {len(junk_emails)} 封垃圾邮件")

        mail.logout()
        
        if all_emails_text:
            result = "\n".join(all_emails_text)
            print(f"🎉 成功获取 {len(all_emails_text)} 封邮件")
        else:
            result = "📭 未找到邮件"
            print("📭 未找到任何邮件")
            
        return result, 200, {'Content-Type': 'text/plain; charset=utf-8'}

    except requests.exceptions.RequestException as e:
        error_msg = f"🌐 网络请求错误: {str(e)}"
        print(error_msg)
        return jsonify({"error": error_msg}), 500
    except imaplib.IMAP4.error as e:
        error_msg = f"📧 邮箱连接或认证错误: {str(e)}"
        print(error_msg)
        return jsonify({"error": error_msg}), 500
    except Exception as e:
        error_msg = f"⚠️ 服务器内部错误: {str(e)}"
        print(error_msg)
        print(traceback.format_exc())
        return jsonify({"error": error_msg}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    print(f"🚀 启动邮件API服务 (Zeabur优化版)...")
    print(f"📍 端口: {port}")
    app.run(host='0.0.0.0', port=port, debug=False)
