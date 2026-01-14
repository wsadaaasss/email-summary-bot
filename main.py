# main.py

# ==============================================================================
# 导入必要的库
# ==============================================================================
import os
import imaplib
import email
from email.header import decode_header
from email.utils import parsedate_to_datetime
from datetime import datetime, timedelta, timezone
import smtplib
from email.mime.text import MIMEText
from email.header import Header
import openai
import json
import markdown2
import re

# ==============================================================================
# 全局常量与配置
# ==============================================================================

SYSTEM_PROMPT = """
# 角色
你是一名专业的学术快讯每日邮件分析助手，任务是根据下方提供的邮件JSON数据，生成一段Markdown格式的摘要报告。

# 任务指令
1. 仔细阅读提供的邮件JSON数据。
2. **不要**生成顶层标题或总览信息，只专注于处理邮件列表。
3. 按顺序逐一处理`{{emails}}`数组中的每一封邮件，并为每封邮件提取以下信息：
   - **发件人**：提取`from_sender`。
   - **主题**：提取`subject`。
   - **摘要**：根据`body_preview`概括核心内容。
4. 严格按照"输出格式要求"生成内容。邮件序号**必须从 {{start_index}} 开始**。

# 文献Alert邮件特殊处理
如果邮件是文献alert（如关注期刊更新、关注特定主题的检索、关注作者最新相关工作、被关注作者或文章的最新引用等），在"摘要"部分必须：
1. 首行说明期刊名称和alert类型
2. 逐条列出每篇论文的详细信息，格式如下：
   - 论文标题
   - 作者
   - 发表日期（如果有）
   - 文章类型（如果有）

# 输出格式要求
#### 邮件 {{start_index}}：[第一封邮件的主题]
- **发件人**：[发件人信息]
- **摘要**：
  [如果是期刊新增alert，按以下格式]
  期刊《Aerospace Science and Technology》新增20篇论文：
  
  1. A Hybrid Deep Learning Framework for Efficient Airfoil Design Optimization
     作者：Abdurrahman Tekin, Tianhang XIAO, Xiongqing YU
     发表日期：11 January 2026
     类型：Research article
  
  2. Noise reduction mechanisms of brush-like trailing-edge extensions on a stalled airfoil
     作者：Zhi Deng, Yong Wang, Zifeng Yang, Donglai Gao, Wen-Li Chen
     发表日期：10 January 2026
     类型：Research article
  
  ...（列出所有论文）
  
# 如果不是文献alert，总结后概括
- **摘要**：[简洁概括]

---

#### 邮件 {{start_index + 1}}：[第二封邮件的主题]
- **发件人**：[发件人信息]
- **摘要**：[简洁概括]
[如果是所关注特定主题的检索、关注作者最新相关工作、被关注作者或文章的最新引用alert，也按以下格式详细列出文章名、作者名]

 1. A Hybrid Deep Learning Framework for Efficient Airfoil Design Optimization
     作者：Abdurrahman Tekin, Tianhang XIAO, Xiongqing YU
     发表日期：11 January 2026
     类型：Research article
  
  2. Noise reduction mechanisms of brush-like trailing-edge extensions on a stalled airfoil
     作者：Zhi Deng, Yong Wang, Zifeng Yang, Donglai Gao, Wen-Li Chen
     发表日期：10 January 2026
     类型：Research article
     
# 如果不是文献alert，总结后概括
- **摘要**：[简洁概括] 


... (以此类推，涉及到论文均按上述标准详细列出文章名、作者等信息直到处理完批次内的所有邮件)

# 特别说明
- 文献alert识别关键词：Web of Science Alert、Google Scholar Alerts、ScienceDirect、Google学术、Alert、New Articles、期刊更新、新论文、Available Online等。

# 待分析的邮件数据
{{emails}}
"""

# 从GitHub Secrets安全地加载环境变量
IMAP_EMAIL = os.environ.get("IMAP_EMAIL")
IMAP_AUTH_CODE = os.environ.get("IMAP_AUTH_CODE")
IMAP_SERVER = os.environ.get("IMAP_SERVER")
IMAP_PORT = int(os.environ.get("IMAP_PORT", 993))
TARGET_FOLDER = os.environ.get("TARGET_FOLDER")
DEEPSEEK_API_KEY = os.environ.get("DEEPSEEK_API_KEY")
SENDER_EMAIL = os.environ.get("SENDER_EMAIL")
SENDER_AUTH_CODE = os.environ.get("SENDER_AUTH_CODE")
RECEIVER_EMAIL = os.environ.get("RECEIVER_EMAIL")
SMTP_SERVER = os.environ.get("SMTP_SERVER")
SMTP_PORT = int(os.environ.get("SMTP_PORT", 587))

# ==============================================================================
# 核心功能函数
# ==============================================================================

def extract_text_from_html(html_content):
    """
    从HTML内容中提取纯文本，保留论文列表的结构。
    """
    if not html_content:
        return ""
    
    # 移除HTML标签，但保留文本内容
    text = re.sub(r'<[^>]+>', '\n', html_content)
    
    # 移除多余的空白行
    text = re.sub(r'\n\s*\n', '\n\n', text)
    
    # 移除HTML实体
    html_entities = {
        '&nbsp;': ' ',
        '&amp;': '&',
        '&lt;': '<',
        '&gt;': '>',
        '&quot;': '"',
        '&#39;': "'"
    }
    for entity, char in html_entities.items():
        text = text.replace(entity, char)
    
    # 清理每行前后的空格
    lines = [line.strip() for line in text.split('\n')]
    text = '\n'.join(lines)
    
    return text.strip()

def get_emails_from_target_date(target_date):
    """
    通过IMAP连接到邮箱，获取指定日期的邮件。
    采用“客户端过滤”策略，并在过滤前将所有邮件时间统一到北京时区，以确保准确性。
    对邮件头和正文的解码增加了容错处理。
    """
    mail_list = []
    beijing_tz = timezone(timedelta(hours=8))

    try:
        conn = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        conn.login(IMAP_EMAIL, IMAP_AUTH_CODE)
        conn.select(f'"{TARGET_FOLDER}"')
        
        fetch_since_dt = target_date - timedelta(days=2)
        fetch_since_str = fetch_since_dt.strftime("%d-%b-%Y")
        search_query = f'(SINCE "{fetch_since_str}")'
        
        status, messages = conn.search(None, search_query)
        if status != "OK":
            print(f"IMAP search failed for query: {search_query}")
            conn.logout()
            return []
            
        email_ids = messages[0].split()
        
        for email_id in reversed(email_ids):
            _, msg_data = conn.fetch(email_id, "(RFC822)")
            msg = email.message_from_bytes(msg_data[0][1])

            try:
                date_header = msg.get("Date")
                if not date_header: continue
                
                email_dt_original = parsedate_to_datetime(date_header)
                
                if email_dt_original.tzinfo is None:
                    email_dt_in_beijing = email_dt_original.replace(tzinfo=timezone.utc).astimezone(beijing_tz)
                else:
                    email_dt_in_beijing = email_dt_original.astimezone(beijing_tz)

                if email_dt_in_beijing.date() != target_date.date():
                    continue

                subject, encoding = decode_header(msg["Subject"])[0]
                if isinstance(subject, bytes):
                    subject = subject.decode(encoding if encoding else "utf-8", errors='ignore')

                from_, encoding = decode_header(msg.get("From"))[0]
                if isinstance(from_, bytes):
                    from_ = from_.decode(encoding if encoding else "utf-8", errors='ignore')

                body = ""
                body_html = ""
                if msg.is_multipart():
                    for part in msg.walk():
                        content_type = part.get_content_type()
                        if content_type == "text/html" and not body_html:
                            try:
                                body_bytes = part.get_payload(decode=True)
                                try: body_html = body_bytes.decode('utf-8')
                                except UnicodeDecodeError:
                                    try: body_html = body_bytes.decode('gbk')
                                    except UnicodeDecodeError: body_html = body_bytes.decode('utf-8', errors='ignore')
                            except: continue
                        elif content_type == "text/plain" and not body:
                            try:
                                body_bytes = part.get_payload(decode=True)
                                try: body = body_bytes.decode('utf-8')
                                except UnicodeDecodeError:
                                    try: body = body_bytes.decode('gbk')
                                    except UnicodeDecodeError: body = body_bytes.decode('utf-8', errors='ignore')
                            except: continue
                else:
                    content_type = msg.get_content_type()
                    try:
                        body_bytes = msg.get_payload(decode=True)
                        try: 
                            if content_type == "text/html":
                                body_html = body_bytes.decode('utf-8')
                            else:
                                body = body_bytes.decode('utf-8')
                        except UnicodeDecodeError:
                            try: 
                                if content_type == "text/html":
                                    body_html = body_bytes.decode('gbk')
                                else:
                                    body = body_bytes.decode('gbk')
                            except UnicodeDecodeError:
                                if content_type == "text/html":
                                    body_html = body_bytes.decode('utf-8', errors='ignore')
                                else:
                                    body = body_bytes.decode('utf-8', errors='ignore')
                    except: body = "无法解码正文。"
                
                # 如果有HTML内容，优先使用HTML提取的文本
                if body_html:
                    body = extract_text_from_html(body_html)
                
                # 增加预览长度到10000字符，以便包含完整的文献列表
                mail_list.append({ "from_sender": from_, "subject": subject, "body_preview": body[:10000] })
            except Exception as e:
                print(f"解析邮件 {email_id.decode()} 时出错: {e}")
                continue
                
        conn.logout()
        print(f"成功从'{TARGET_FOLDER}'文件夹获取并过滤出 {len(mail_list)} 封邮件。")
        return mail_list
    except Exception as e:
        print(f"获取邮件失败: {e}")
        return []

def summarize_single_batch(client, email_batch, start_index):
    """
    【辅助函数】使用统一的SYSTEM_PROMPT，处理单批次的邮件。
    """
    emails_json_str = json.dumps(email_batch, ensure_ascii=False, indent=2)
    
    prompt_filled = SYSTEM_PROMPT.replace("{{emails}}", emails_json_str)
    prompt_filled = prompt_filled.replace("{{start_index}}", str(start_index))
    
    try:
        response = client.chat.completions.create(
            model="deepseek-chat",
            messages=[{"role": "system", "content": prompt_filled}]
        )
        return response.choices[0].message.content
    except Exception as e:
        print(f"处理批次 (起始序号 {start_index}) 时调用API失败: {e}")
        return f"--- \n\n#### 处理邮件 {start_index} 到 {start_index + len(email_batch) - 1} 时出错\n- **错误详情**: `{e}`\n\n---"


def summarize_with_llm(email_list, batch_size=5):
    """
    协调分批处理邮件列表的总结任务。
    """
    if not email_list:
        return "### 每日邮件汇总\n**总览：共 0 封邮件**\n\n--- \n\n今日没有收到新邮件。"
        
    client = openai.OpenAI(
        api_key=DEEPSEEK_API_KEY,
        base_url="https://api.deepseek.com/v1"
    )
    
    total_emails = len(email_list)
    report_parts = [f"### 每日邮件汇总\n**总览：共 {total_emails} 封邮件**\n\n---"]
    
    print(f"开始分批总结 {total_emails} 封邮件，每批最多 {batch_size} 封...")
    
    for i in range(0, total_emails, batch_size):
        batch = email_list[i:i + batch_size]
        start_index = i + 1
        print(f"  正在处理第 {i//batch_size + 1} 批 (邮件 {start_index} 到 {min(i+batch_size, total_emails)})...")
        
        batch_summary = summarize_single_batch(client, batch, start_index)
        report_parts.append(batch_summary)

    return "\n".join(report_parts)


def send_email_notification(summary_md, date_for_subject):
    """
    将Markdown报告转换为HTML并通过SMTP (STARTTLS) 发送。
    """
    if not SENDER_EMAIL or not SENDER_AUTH_CODE or not RECEIVER_EMAIL:
        print("发送邮件所需的环境变量不完整，跳过发送。")
        return

    html_content = markdown2.markdown(summary_md, extras=["tables", "fenced-code-blocks"])
    message = MIMEText(html_content, 'html', 'utf-8')
    
    subject_str = f"每日邮件总结 - {date_for_subject.strftime('%Y-%m-%d')}"
    message['Subject'] = Header(subject_str, 'utf-8')
    message['From'] = SENDER_EMAIL
    message['To'] = RECEIVER_EMAIL

    try:
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=30)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_AUTH_CODE)
        server.sendmail(SENDER_EMAIL, [RECEIVER_EMAIL], message.as_string())
        server.quit()
        print(f"成功发送邮件总结到 {RECEIVER_EMAIL}！")
    except Exception as e:
        print(f"发送邮件失败: {e}")

# ==============================================================================
# 主执行入口
# ==============================================================================
if __name__ == "__main__":
    required_vars = ["IMAP_EMAIL", "IMAP_AUTH_CODE", "IMAP_SERVER", "IMAP_PORT", "TARGET_FOLDER", "DEEPSEEK_API_KEY", 
                     "SENDER_EMAIL", "SENDER_AUTH_CODE", "RECEIVER_EMAIL", "SMTP_SERVER", "SMTP_PORT"]
    if not all(os.environ.get(var) for var in required_vars):
        print("错误：一个或多个必要的环境变量未设置。")
        exit(1)

    print(f"任务启动于 (UTC): {datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')}")
    
    beijing_timezone = timezone(timedelta(hours=8))
    # 使用带时区的 now() 以确保夏令时等边缘情况的准确性
    beijing_now = datetime.now(beijing_timezone)
    
    target_day = beijing_now - timedelta(days=1)
    print(f"将要总结的日期是 (北京时间): {target_day.strftime('%Y-%m-%d')}")
    
    emails = get_emails_from_target_date(target_day)
    summary_report = summarize_with_llm(emails)
    send_email_notification(summary_report, target_day)
    
    print(f"任务执行完毕于 (UTC): {datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')}")
