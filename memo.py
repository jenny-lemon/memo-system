# -*- coding: utf-8 -*-
import requests
import re
from datetime import datetime

BASE_URL = "https://backend.lemonclean.com.tw"

RUNTIME_EMAIL = ""
RUNTIME_PASSWORD = ""


# ========================
# 工具
# ========================
def set_runtime_credentials(email, password):
    global RUNTIME_EMAIL, RUNTIME_PASSWORD
    RUNTIME_EMAIL = email
    RUNTIME_PASSWORD = password


def parse_date(s):
    try:
        return datetime.strptime(s.replace("-", "/"), "%Y/%m/%d")
    except:
        return None


def log_safe(log, msg):
    if log:
        log(msg)


# ========================
# 登入
# ========================
def login(session):
    session.get(f"{BASE_URL}/login")
    session.post(f"{BASE_URL}/login", data={
        "email": RUNTIME_EMAIL,
        "password": RUNTIME_PASSWORD
    })


# ========================
# 解析列表
# ========================
def parse_list(html):
    rows = []
    lines = html.splitlines()

    for line in lines:
        if "LC00" not in line:
            continue

        order = re.search(r'(LC\d+)', line)
        date = re.search(r'(\d{4}/\d{2}/\d{2})', line)

        if not order:
            continue

        rows.append({
            "order_no": order.group(1),
            "date": date.group(1) if date else "",
            "line": line
        })

    return rows


# ========================
# 訂單詳情
# ========================
def get_detail(session, order_no):
    url = f"{BASE_URL}/purchase/edit/{order_no.replace('LC','')}"
    res = session.get(url)
    text = res.text

    phone = re.search(r'09\d{8}', text)
    address = re.search(r'台.+?\d+樓', text)
    name = re.search(r'客戶.*?>(.*?)<', text)
    date = re.search(r'(\d{4}-\d{2}-\d{2})', text)

    return {
        "phone": phone.group(0) if phone else "",
        "address": address.group(0) if address else "",
        "name": name.group(1) if name else "",
        "date": date.group(1).replace("-", "/") if date else "",
        "date_obj": parse_date(date.group(1)) if date else None,
        "html": text
    }


# ========================
# 找上一筆
# ========================
def find_previous(session, phone, address, current_date):

    url = f"{BASE_URL}/purchase?phone={phone}&purchase_status=1"
    res = session.get(url)

    items = parse_list(res.text)

    best = None

    for i in items:
        d = get_detail(session, i["order_no"])

        if d["address"] != address:
            continue

        if not d["date_obj"]:
            continue

        if d["date_obj"] >= current_date:
            continue

        if "已處理" not in d["html"]:
            continue

        if not best or d["date_obj"] > best["date_obj"]:
            best = d
            best["order_no"] = i["order_no"]

    return best


# ========================
# 更新訂單
# ========================
def update(session, order_no, memo):
    url = f"{BASE_URL}/purchase?id={order_no.replace('LC','')}"

    session.post(url, data={
        "progress_status": "1",
        "notice": memo
    })

    check = session.get(f"{BASE_URL}/purchase/edit/{order_no.replace('LC','')}").text

    return ("已處理" in check) and (memo[:5] in check)


# ========================
# 單筆處理
# ========================
def process_one(session, order_no, log):

    d = get_detail(session, order_no)

    log_safe(log, f"訂單: {order_no}")
    log_safe(log, f"客戶: {d['name']}")
    log_safe(log, f"電話: {d['phone']}")
    log_safe(log, f"地址: {d['address']}")
    log_safe(log, f"日期: {d['date']}")

    prev = find_previous(session, d["phone"], d["address"], d["date_obj"])

    if not prev:
        log_safe(log, f"❌ 沒有上一筆")
        return False

    log_safe(log, f"[上一筆] {prev['date']} {prev['order_no']}")

    memo_text = f"沿用上一筆 {prev['order_no']}"

    ok = update(session, order_no, memo_text)

    if ok:
        log_safe(log, f"✅ 成功 {order_no}")
    else:
        log_safe(log, f"❌ 驗證失敗 {order_no}")

    return ok


# ========================
# By 電話
# ========================
def main_by_phone(phone, ui_logger=None):

    session = requests.Session()
    login(session)

    html = session.get(f"{BASE_URL}/purchase?phone={phone}&purchase_status=1").text
    items = parse_list(html)

    success = 0

    for i in items:
        ok = process_one(session, i["order_no"], ui_logger)
        if ok:
            success += 1

    return {
        "processed": len(items),
        "success": success,
        "failed": len(items) - success,
        "updated_orders": success
    }


# ========================
# ⭐ By 搜尋條件（你缺的）
# ========================
def main_by_conditions(date_s, limit=
