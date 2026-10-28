#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Alo Yoga 邮件监控 - GitHub Actions 云运行版
不需要本地电脑开机，GitHub 免费帮你跑
"""

import imaplib
import email
import json
import os
import re
import urllib.request
import urllib.parse
import base64
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed

# 从环境变量读取配置
EMAIL_LIST_RAW = os.environ.get("EMAIL_LIST", "")
BARK_URL = os.environ.get("BARK_URL", "")

# 成功/失败关键词
# 只有 received 才算成功，只有 canceled 才算失败，其他都继续检测
SUCCESS_KEYWORDS = ["Payment for", "has been received"]
FAIL_KEYWORDS = ["has been canceled"]

IMAP_SERVER = "outlook.office365.com"
IMAP_PORT = 993
TOKEN_URL = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
SCOPE = "https://outlook.office.com/IMAP.AccessAsUser.All"

# 状态文件（用 GitHub Secrets 存储，运行时读取）
STATE_JSON = os.environ.get("STATE_JSON", "{}")

def parse_email_line(line):
    parts = line.strip().split("----")
    if len(parts) < 5:
        return None
    return {
        "email": parts[0].strip(),
        "password": parts[1].strip(),
        "client_id": parts[2].strip(),
        "refresh_token": parts[3].strip(),
        "group": parts[4].strip(),
    }

def load_email_list():
    emails = []
    for line in EMAIL_LIST_RAW.split("\n"):
        line = line.strip()
        if not line:
            continue
        info = parse_email_line(line)
        if info:
            emails.append(info)
    return emails

def load_state():
    try:
        return json.loads(STATE_JSON)
    except:
        return {"completed": {}, "last_check": None}

def get_access_token(client_id, refresh_token):
    data = urllib.parse.urlencode({
        "grant_type": "refresh_token",
        "refresh_token": refresh_token,
        "client_id": client_id,
        "scope": SCOPE,
    }).encode("utf-8")
    req = urllib.request.Request(TOKEN_URL, data=data)
    with urllib.request.urlopen(req, timeout=30) as resp:
        result = json.loads(resp.read().decode("utf-8"))
    return result.get("access_token")

def imap_auth(imap, email_addr, access_token):
    auth_string = f"user={email_addr}\x01auth=Bearer {access_token}\x01\x01"
    imap.authenticate("XOAUTH2", lambda x: auth_string.encode("utf-8"))

def check_email(email_info, state):
    email_addr = email_info["email"]
    if email_addr in state.get("completed", {}):
        return "already_done", ""

    try:
        access_token = get_access_token(email_info["client_id"], email_info["refresh_token"])
        if not access_token:
            return "error", "获取 access_token 失败"

        imap = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        imap_auth(imap, email_addr, access_token)
        imap.select("INBOX")

        since_date = (datetime.now() - timedelta(days=7)).strftime("%d-%b-%Y")
        status, messages = imap.search(None, f'(SINCE "{since_date}")')

        if status != "OK":
            imap.logout()
            return "error", "搜索邮件失败"

        msg_ids = messages[0].split()
        if not msg_ids:
            imap.logout()
            return "pending", "最近7天无邮件"

        found_result = None
        for msg_id in reversed(msg_ids[-20:]):
            status, msg_data = imap.fetch(msg_id, "(BODY[HEADER.FIELDS (SUBJECT FROM DATE)])")
            if status != "OK":
                continue
            raw_headers = msg_data[0][1]
            msg = email.message_from_bytes(raw_headers)
            subject = str(email.header.make_header(email.header.decode_header(msg["Subject"] or "")))
            sender = str(email.header.make_header(email.header.decode_header(msg["From"])))

            if "alo yoga" not in sender.lower() and "alo yoga" not in subject.lower():
                continue

            if all(kw.lower() in subject.lower() for kw in SUCCESS_KEYWORDS):
                found_result = ("success", f"✅ 成功 — {subject}")
                break
            if any(kw.lower() in subject.lower() for kw in FAIL_KEYWORDS):
                found_result = ("fail", f"❌ 失败 — {subject}")
                break

        imap.logout()

        if found_result:
            return found_result[0], found_result[1]
        else:
            return "pending", "无成功/失败结果"

    except Exception as e:
        return "error", f"异常: {str(e)[:100]}"

def bark_push(title, message):
    if not BARK_URL:
        print("Bark URL 未配置，跳过推送")
        return
    try:
        url = f"{BARK_URL}/{urllib.parse.quote(title)}/{urllib.parse.quote(message)}"
        urllib.request.urlopen(url, timeout=10)
        print("Bark 推送成功")
    except Exception as e:
        print(f"Bark 推送失败: {e}")

def main():
    print(f"=== Alo Yoga 监控开始 {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===")

    state = load_state()
    emails = load_email_list()
    pending_emails = [e for e in emails if e["email"] not in state.get("completed", {})]
    
    print(f"共 {len(emails)} 个邮箱，已完成 {len(state.get('completed', {}))} 个，待检测 {len(pending_emails)} 个")

    import threading
    lock = threading.Lock()
    new_results = []
    completed_count = [0]

    def check_one(info):
        email_addr = info["email"]
        status, detail = check_email(info, state)
        with lock:
            completed_count[0] += 1
            print(f"[{completed_count[0]}/{len(pending_emails)}] {email_addr}: {status}")
        return email_addr, info, status, detail

    with ThreadPoolExecutor(max_workers=10) as executor:
        futures = [executor.submit(check_one, info) for info in pending_emails]
        for future in as_completed(futures):
            email_addr, info, status, detail = future.result()
            if status in ("success", "fail"):
                with lock:
                    state.setdefault("completed", {})[email_addr] = {
                        "result": status,
                        "detail": detail,
                        "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "group": info["group"],
                    }
                    new_results.append((email_addr, status, detail))

    state["last_check"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # 输出新的状态 JSON（供 Actions 保存）
    print("===STATE_JSON_START===")
    print(json.dumps(state, ensure_ascii=False))
    print("===STATE_JSON_END===")

    if new_results:
        print(f"\n发现 {len(new_results)} 个新结果！")
        summary_lines = []
        for addr, status, detail in new_results:
            line = f"{addr}: {detail}"
            summary_lines.append(line)
            print(line)

        full_message = "\n".join(summary_lines)
        title = f"Alo Yoga 监控 — {len(new_results)} 个新结果"
        bark_push(title, full_message)
    else:
        print("\n本次无新结果。")

    print(f"=== 结束 {datetime.now().strftime('%H:%M:%S')} ===")

if __name__ == "__main__":
    main()
