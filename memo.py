import requests
import re
from datetime import datetime

BASE_URL = "https://backend.lemonclean.com.tw"


# =========================
# 工具
# =========================

def parse_date(s):
    if not s:
        return None
    s = s.replace("-", "/")
    try:
        return datetime.strptime(s, "%Y/%m/%d")
    except:
        return None


def extract_service_date(text):
    text = str(text or "")

    m = re.search(r'(\d{4}[/-]\d{2}[/-]\d{2})\s*\([一二三四五六日]\)', text)
    if m:
        return m.group(1).replace("-", "/")

    m = re.search(r'(\d{4}[/-]\d{2}[/-]\d{2})', text)
    if m:
        return m.group(1).replace("-", "/")

    return ""


def normalize_address(addr):
    if not addr:
        return ""
    return addr.replace("之一", "之1").replace("之二", "之2").strip()


# =========================
# 登入
# =========================

def login(session, email, password):
    session.get(f"{BASE_URL}/login")
    data = {
        "email": email,
        "password": password
    }
    session.post(f"{BASE_URL}/login", data=data)


# =========================
# 列表解析
# =========================

def parse_list(html):
    rows = []
    lines = html.splitlines()

    for line in lines:
        if "LC00" not in line:
            continue

        order = re.search(r'(LC\d+)', line)
        date = re.search(r'(\d{4}/\d{2}/\d{2})', line)
        phone = re.search(r'09\d{8}', line)

        if not order:
            continue

        rows.append({
            "order_no": order.group(1),
            "raw_date": date.group(1) if date else "",
            "phone": phone.group(0) if phone else "",
            "line": line
        })

    return rows


# =========================
# 訂單詳情
# =========================

def get_order_detail(session, order_no):
    url = f"{BASE_URL}/purchase/edit/{order_no.replace('LC','')}"
    res = session.get(url)
    text = res.text

    phone = re.search(r'09\d{8}', text)
    address = re.search(r'台[北中南].+?\d+樓', text)
    name = re.search(r'客戶.*?>(.*?)<', text)

    service_date = extract_service_date(text)

    return {
        "phone": phone.group(0) if phone else "",
        "address": normalize_address(address.group(0) if address else ""),
        "name": name.group(1) if name else "",
        "service_date": service_date,
        "service_date_obj": parse_date(service_date),
        "html": text
    }


# =========================
# 找上一筆
# =========================

def find_previous(session, phone, address, current_date):

    url = f"{BASE_URL}/purchase?phone={phone}&purchase_status=1"
    res = session.get(url)
    items = parse_list(res.text)

    candidates = []

    for i in items:
        d = get_order_detail(session, i["order_no"])

        if d["address"] != address:
            continue

        if not d["service_date_obj"]:
            continue

        if d["service_date_obj"] >= current_date:
            continue

        # 必須已處理
        if "已處理" not in d["html"]:
            continue

        candidates.append((i["order_no"], d))

    candidates.sort(key=lambda x: x[1]["service_date_obj"], reverse=True)

    return candidates[0] if candidates else None


# =========================
# 更新訂單
# =========================

def update_order(session, order_no, memo):

    url = f"{BASE_URL}/purchase?id={order_no.replace('LC','')}"

    data = {
        "progress_status": "1",
        "notice": memo
    }

    session.post(url, data=data)

    # 驗證
    check = session.get(f"{BASE_URL}/purchase/edit/{order_no.replace('LC','')}").text

    success = ("已處理" in check) and (memo[:10] in check)

    return success


# =========================
# 主流程（單筆）
# =========================

def process_order(session, order_no, log):

    d = get_order_detail(session, order_no)

    log(f"訂單: {order_no}")
    log(f"客戶: {d['name']}")
    log(f"電話: {d['phone']}")
    log(f"地址: {d['address']}")
    log(f"日期: {d['service_date']}")

    prev = find_previous(session, d["phone"], d["address"], d["service_date_obj"])

    if not prev:
        log(f"❌ 處理失敗 {order_no}：沒有上一筆")
        return 0, False

    prev_no, prev_data = prev

    log(f"[上一筆] {prev_data['service_date']} {prev_no}")

    memo = "自動帶入：" + prev_data["name"]

    success = update_order(session, order_no, memo)

    if success:
        log(f"✅ 驗證成功 {order_no}")
        return 1, True
    else:
        log(f"❌ 驗證失敗 {order_no}")
        return 0, False
