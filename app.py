from flask import Flask, request, jsonify
import requests
import imaplib
import email
from bs4 import BeautifulSoup
from email.header import decode_header
import chardet
import os
import traceback
import re

# 环境变量
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

def matches_subject_filter(subject, subject_filter):
    """检查邮件主题是否匹配筛选条件"""
    if not subject_filter:
        return True  # 如果没有筛选条件，返回所有邮件
    
    # 将主题和筛选条件都转为小写进行比较（不区分大小写）
    subject_lower = subject.lower()
    filter_lower = subject_filter.lower()
    
    # 支持多种匹配方式
    # 1. 直接包含匹配
    if filter_lower in subject_lower:
        return True
    
    # 2. 支持逗号分隔的多个关键词（任意一个匹配即可）
    if ',' in filter_lower:
        keywords = [kw.strip() for kw in filter_lower.split(',')]
        for keyword in keywords:
            if keyword and keyword in subject_lower:
                return True
    
    # 3. 支持空格分隔的多个关键词（所有关键词都必须包含）
    elif ' ' in filter_lower:
        keywords = filter_lower.split()
        return all(keyword in subject_lower for keyword in keywords if keyword)
    
    return False

def get_emails_from_folder(mail, folder_name, max_emails=20, subject_filter=None):
    """从指定文件夹获取邮件内容，支持主题筛选"""
    emails_data = []
    try:
        status, messages = mail.select(folder_name, readonly=True)
        if status != "OK":
            print(f"选择文件夹 {folder_name} 失败: {status}")
            return emails_data

        status, message_ids = mail.search(None, 'ALL')
        if status != "OK":
            print(f"邮件搜索失败: {status}")
            return emails_data

        message_list = message_ids[0].split()
        # 如果有筛选条件，增加搜索范围以找到更多匹配的邮件
        search_limit = max_emails * 3 if subject_filter else max_emails
        limited_messages = message_list[-search_limit:] if len(message_list) > search_limit else message_list

        matched_count = 0
        for i, message_id in enumerate(limited_messages):
            # 如果已经找到足够的匹配邮件，停止搜索
            if matched_count >= max_emails:
                break
                
            try:
                status, msg_data = mail.fetch(message_id, '(RFC822)')
                if status != "OK":
                    continue

                for response_part in msg_data:
                    if isinstance(response_part, tuple):
                        msg = email.message_from_bytes(response_part[1])
                        subject = decode_mime_words(msg.get("subject", ""))
                        
                        # 应用主题筛选
                        if not matches_subject_filter(subject, subject_filter):
                            continue  # 跳过不匹配的邮件
                        
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
                            f"邮件编号: {matched_count + 1}\n" \
                            f"邮件主题: {subject}\n" \
                            f"收件时间: {date}\n" \
                            f"发件人: {sender}\n" \
                            f"邮件正文:\n{body[:800]}{'...' if len(body) > 800 else ''}\n" \
                            f"\n--------------------------------------------------\n"
                        )
                        matched_count += 1
                        break  # 处理完这封邮件，继续下一封
                        
            except Exception as e:
                print(f"处理邮件时出错: {str(e)}")
                continue

    except Exception as e:
        print(f"获取邮件文件夹 {folder_name} 时出错: {str(e)}")
    
    return emails_data

@app.route('/')
def index():
    """根路径，提供API说明"""
    return """
    <html>
    <head>
        <title>邮件获取API - 增强版</title>
        <meta charset="UTF-8">
        <style>
            body { font-family: Arial, sans-serif; max-width: 900px; margin: 0 auto; padding: 20px; background: #f5f5f5; }
            .container { background: white; padding: 30px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
            h1 { color: #2c3e50; text-align: center; }
            .status { background: #d4edda; border: 1px solid #c3e6cb; padding: 15px; border-radius: 5px; margin: 20px 0; }
            .api-info { background: #e3f2fd; border: 1px solid #bbdefb; padding: 20px; border-radius: 5px; margin: 20px 0; }
            .feature { background: #fff3cd; border: 1px solid #ffeaa7; padding: 15px; border-radius: 5px; margin: 20px 0; }
            code { background: #f8f9fa; padding: 5px 10px; border-radius: 3px; font-family: monospace; }
            .example { background: #f8f9fa; padding: 15px; border-radius: 5px; margin: 10px 0; }
        </style>
    </head>
    <body>
        <div class="container">
            <h1>📧 邮件获取API - 增强版</h1>
            <div class="status">
                <h2>🎉 部署成功！</h2>
                <p>服务运行正常，国内访问优化</p>
                <p><strong>部署平台：</strong> Zeabur</p>
                <p><strong>新功能：</strong> 支持邮件主题筛选 🔍</p>
            </div>
            
            <div class="feature">
                <h3>🆕 新增功能：主题筛选</h3>
                <p>现在可以通过 <code>subject_filter</code> 参数筛选特定主题的邮件！</p>
            </div>
            
            <div class="api-info">
                <h3>📡 API使用方法</h3>
                
                <h4>基本用法（获取所有邮件）：</h4>
                <div class="example">
                    <code>GET /get_emails?email_address=你的邮箱&refresh_token=你的令牌</code>
                </div>
                
                <h4>筛选特定主题（新功能）：</h4>
                <div class="example">
                    <code>GET /get_emails?email_address=你的邮箱&refresh_token=你的令牌&subject_filter=Magic</code>
                </div>
                
                <h4>参数说明：</h4>
                <ul>
                    <li><strong>email_address</strong>: 你的邮箱地址（必需）</li>
                    <li><strong>refresh_token</strong>: 你的刷新令牌（必需）</li>
                    <li><strong>subject_filter</strong>: 主题筛选关键词（可选）</li>
                    <li><strong>max_emails</strong>: 最大返回邮件数量（可选，默认10）</li>
                </ul>
                
                <h4>筛选示例：</h4>
                <div class="example">
                    <p><strong>包含"Magic"的邮件：</strong></p>
                    <code>subject_filter=Magic</code>
                    
                    <p><strong>包含"验证码"的邮件：</strong></p>
                    <code>subject_filter=验证码</code>
                    
                    <p><strong>包含多个关键词之一（任意匹配）：</strong></p>
                    <code>subject_filter=Magic,验证码,Verification</code>
                    
                    <p><strong>同时包含多个关键词（全部匹配）：</strong></p>
                    <code>subject_filter=Magic Email Verification</code>
                </div>
                
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
        "version": "2.0",
        "features": ["subject_filter", "max_emails"]
    })

@app.route('/get_emails')
def get_emails_api():
    try:
        email_address = request.args.get('email_address')
        refresh_token = request.args.get('refresh_token')
        subject_filter = request.args.get('subject_filter')  # 新增：主题筛选参数
        max_emails = int(request.args.get('max_emails', 10))  # 新增：最大邮件数量参数

        print(f"📧 收到请求 - 邮箱: {email_address}")
        if subject_filter:
            print(f"🔍 主题筛选: {subject_filter}")
        print(f"📊 最大邮件数: {max_emails}")

        if not email_address or not refresh_token:
            return jsonify({"error": "❌ 缺少必要参数: email_address 或 refresh_token"}), 400

        # 验证max_emails参数
        if max_emails < 1 or max_emails > 50:
            return jsonify({"error": "❌ max_emails 参数必须在1-50之间"}), 400

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
        inbox_emails = get_emails_from_folder(mail, "INBOX", max_emails, subject_filter)
        all_emails_text.extend(inbox_emails)
        print(f"✅ 获取到 {len(inbox_emails)} 封收件箱邮件")

        # 如果收件箱没有找到足够的邮件，继续搜索垃圾邮件文件夹
        remaining_emails = max_emails - len(inbox_emails)
        if remaining_emails > 0:
            print("🗑️ 正在获取垃圾邮件...")
            junk_emails = get_emails_from_folder(mail, "Junk", remaining_emails, subject_filter)
            if not junk_emails:
                junk_emails = get_emails_from_folder(mail, "Junk Email", remaining_emails, subject_filter)
            all_emails_text.extend(junk_emails)
            print(f"✅ 获取到 {len(junk_emails)} 封垃圾邮件")

        mail.logout()
        
        if all_emails_text:
            result = "\n".join(all_emails_text)
            total_found = len(all_emails_text)
            filter_info = f" (筛选条件: {subject_filter})" if subject_filter else ""
            print(f"🎉 成功获取 {total_found} 封邮件{filter_info}")
            
            # 在结果开头添加统计信息
            stats = f"=== 邮件获取统计 ===\n"
            stats += f"总计找到: {total_found} 封邮件\n"
            if subject_filter:
                stats += f"筛选条件: {subject_filter}\n"
            stats += f"获取时间: {email.utils.formatdate()}\n"
            stats += f"========================\n\n"
            
            result = stats + result
        else:
            filter_info = f"（筛选条件: {subject_filter}）" if subject_filter else ""
            result = f"📭 未找到邮件{filter_info}"
            print(f"📭 未找到任何邮件{filter_info}")
            
        return result, 200, {'Content-Type': 'text/plain; charset=utf-8'}

    except ValueError as e:
        error_msg = f"❌ 参数错误: {str(e)}"
        print(error_msg)
        return jsonify({"error": error_msg}), 400
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
    print(f"🚀 启动邮件API服务（增强版）...")
    print(f"📍 端口: {port}")
    print(f"🆕 新功能: 主题筛选支持")
    app.run(host='0.0.0.0', port=port, debug=False)
