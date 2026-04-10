# -*- coding: utf-8 -*-
import re
import time
from datetime import datetime
from typing import Optional, List, Dict, Callable
from urllib.parse import urljoin, urlparse, parse_qs

import requests
from bs4 import BeautifulSoup
import gspread
from google.oauth2.service_account import Credentials

import env

try:
    import streamlit as st
except Exception:
    st = None


# ========================
# ENV
# ========================
ENV_NAME = getattr(env, "ENV", "prod").lower()
BASE_URL_DEV = getattr(env, "BASE_URL_DEV", "https://backend-dev.lemonclean.com.tw")
BASE_URL_PROD = getattr(env, "BASE_URL_PROD", "https://backend.lemonclean.com.tw")

BASE_URL = ""
LOGIN_URL = ""
PURCHASE_URL = ""


def set_env(env_name: str):
    global ENV_NAME, BASE_URL, LOGIN_URL, PURCHASE_URL
    ENV_NAME = (env_name or "prod").lower()
    BASE_URL = BASE_URL_DEV if ENV_NAME == "dev" else BASE_URL_PROD
    BASE_URL = BASE_URL.rstrip("/")
    LOGIN_URL = f"{BASE_URL}/login"
    PURCHASE_URL = f"{BASE_URL}/purchase"


set_env(ENV_NAME)

# ========================
# Runtime credentials
# ========================
RUNTIME_EMAIL = ""
RUNTIME_PASSWORD = ""


def set_runtime_credentials(email: str, password: str):
    global RUNTIME_EMAIL, RUNTIME_PASSWORD
    RUNTIME_EMAIL = (email or "").strip()
    RUNTIME_PASSWORD = (password or "").strip()


# ========================
# Settings
# ========================
WORKSHEET_NAME = getattr(env, "WORKSHEET_NAME", "memo")
LOG_SHEET_NAME = getattr(env, "LOG_SHEET_NAME", "memo_log")
GOOGLE_SERVICE_ACCOUNT_FILE = getattr(env, "GOOGLE_SERVICE_ACCOUNT_FILE", "")
SLEEP_SECONDS = float(getattr(env, "SLEEP_SECONDS", 0.5))

REQUEST_TIMEOUT = 30
MAX_RETRIES = 3
RETRY_BACKOFF = 1.2

CURRENT_ROW_LOGS: List[str] = []


# ========================
# Logger
# ========================
def make_logger(ui_logger: Optional[Callable[[str], None]] = None):
    def _log(msg: str):
        msg = str(msg)
        print(msg, flush=True)
        CURRENT_ROW_LOGS.append(msg)
        if ui_logger:
            ui_logger(msg)
    return _log


# ========================
# Helpers
# ========================
def with_retry(fn, *args, **kwargs):
    last_err = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            return fn(*args, **kwargs)
        except Exception as e:
            last_err = e
            if attempt >= MAX_RETRIES:
                break
            time.sleep(RETRY_BACKOFF * attempt)
    raise last_err


def session_get(session: requests.Session, url: str, **kwargs):
    return with_retry(session.get, url, timeout=REQUEST_TIMEOUT, **kwargs)


def session_post(session: requests.Session, url: str, **kwargs):
    return with_retry(session.post, url, timeout=REQUEST_TIMEOUT, **kwargs)


def normalize_phone(p: str) -> str:
    return re.sub(r"\D+", "", str(p or ""))


def normalize_text(t: str) -> str:
    return re.sub(r"\s+", "", str(t or ""))


def normalize_address(addr: str) -> str:
    # 地址需完全相同，只去空白
    return normalize_text(addr)


def same_address(a: str, b: str) -> bool:
    na = normalize_address(a)
    nb = normalize_address(b)
    return bool(na and nb and na == nb)


def clip_text(text: str, limit: int = 50000) -> str:
    return str(text or "")[:limit]


def safe_cell(row: List[str], idx_1_based: int) -> str:
    i = idx_1_based - 1
    return str(row[i]).strip() if i < len(row) else ""


def parse_date(t: str):
    if not t:
        return None

    s = str(t).strip()
    for fmt in [
        "%Y/%m/%d",
        "%Y-%m-%d",
        "%Y/%m/%d %H:%M:%S",
        "%Y-%m-%d %H:%M:%S",
        "%Y/%m/%d %H:%M",
        "%Y-%m-%d %H:%M",
    ]:
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            pass

    m = re.search(r"(\d{4})[/-](\d{1,2})[/-](\d{1,2})", s)
    if m:
        y, mo, d = map(int, m.groups())
        return datetime(y, mo, d)
    return None


def parse_row_spec(spec: str) -> List[int]:
    rows = set()
    for p in str(spec).split(","):
        p = p.strip()
        if not p:
            continue
        if "-" in p:
            a, b = map(int, p.split("-", 1))
            if a > b:
                a, b = b, a
            rows.update(range(a, b + 1))
        else:
            rows.add(int(p))
    return sorted(x for x in rows if x >= 2)


def build_absolute_url(href: str) -> str:
    return urljoin(BASE_URL + "/", str(href or "").strip())


def parse_query_params_from_url(url: str) -> Dict[str, str]:
    parsed = urlparse(url)
    qs = parse_qs(parsed.query)
    result = {}
    for k, v in qs.items():
        if v:
            result[k] = v[0]
    return result


def extract_address_from_text_block(text: str) -> str:
    lines = [x.strip() for x in str(text or "").splitlines() if x.strip()]
    for i, line in enumerate(lines):
        if any(k in line for k in ["市", "縣", "區", "鄉", "鎮", "路", "街", "段", "巷", "弄", "號"]):
            if "付款" in line or "服務狀態" in line or "付款狀態" in line:
                continue
            addr = line
            if i + 1 < len(lines):
                nxt = lines[i + 1]
                if any(k in nxt for k in ["樓", "室", "之", "A", "B", "C", "D", "1", "2", "3", "4", "5"]):
                    addr += nxt
            return addr
    return ""


def extract_name_from_text_block(text: str) -> str:
    lines = [x.strip() for x in str(text or "").splitlines() if x.strip()]
    for line in lines:
        if re.search(r"^[\u4e00-\u9fff]{2,4}$", line):
            return line
    return ""


# ========================
# Google Sheets
# ========================
def get_spreadsheet():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    if st is not None:
        try:
            creds = Credentials.from_service_account_info(
                dict(st.secrets["GOOGLE_SERVICE_ACCOUNT"]),
                scopes=scopes,
            )
            gc = gspread.authorize(creds)
            return gc.open_by_key(env.SHEET_ID)
        except Exception:
            pass

    creds = Credentials.from_service_account_file(
        GOOGLE_SERVICE_ACCOUNT_FILE,
        scopes=scopes,
    )
    gc = gspread.authorize(creds)
    return gc.open_by_key(env.SHEET_ID)


def get_ws():
    return get_spreadsheet().worksheet(WORKSHEET_NAME)


def get_log_ws():
    sh = get_spreadsheet()
    try:
        return sh.worksheet(LOG_SHEET_NAME)
    except Exception:
        ws = sh.add_worksheet(title=LOG_SHEET_NAME, rows=1000, cols=20)
        ws.append_row([
            "執行時間",
            "來源",
            "查詢值",
            "電話",
            "客戶姓名",
            "地址",
            "目前訂單",
            "目前日期",
            "前次訂單",
            "前次日期",
            "前次客服備註",
            "回寫筆數",
            "狀態",
            "錯誤訊息",
            "完整LOG",
        ])
        return ws


def append_log_row(
    log_ws,
    source_type: str,
    source_value: str,
    phone: str,
    name: str,
    address: str,
    current_order: str,
    current_date: str,
    prev_order: str,
    prev_date: str,
    prev_notice: str,
    updated_orders: int,
    status: str,
    error_msg: str,
    full_log: str,
):
    with_retry(
        log_ws.append_row,
        [
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            source_type,
            source_value,
            phone,
            name,
            address,
            current_order,
            current_date,
            prev_order,
            prev_date,
            clip_text(prev_notice, 2000),
            updated_orders,
            status,
            error_msg,
            clip_text(full_log, 20000),
        ],
    )


def apply_sheet_presentation(ws, updated_rows: List[int]):
    if not updated_rows:
        return

    sheet_id = ws._properties["sheetId"]
    requests_body = []

    for row_num in updated_rows:
        requests_body.append({
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "ROWS",
                    "startIndex": row_num - 1,
                    "endIndex": row_num,
                },
                "properties": {"pixelSize": 21},
                "fields": "pixelSize"
            }
        })

    requests_body.append({
        "repeatCell": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": 1,
                "startColumnIndex": 22,  # W
                "endColumnIndex": 24,    # X 後一欄
            },
            "cell": {
                "userEnteredFormat": {
                    "wrapStrategy": "CLIP",
                    "verticalAlignment": "MIDDLE"
                }
            },
            "fields": "userEnteredFormat.wrapStrategy,userEnteredFormat.verticalAlignment"
        }
    })

    with_retry(ws.spreadsheet.batch_update, {"requests": requests_body})


# ========================
# Backend login
# ========================
def login():
    email = RUNTIME_EMAIL
    password = RUNTIME_PASSWORD

    if not email or not password:
        raise RuntimeError("缺少 Email / Password")

    s = requests.Session()
    s.headers.update({
        "User-Agent": "Mozilla/5.0",
        "Accept-Language": "zh-TW,zh;q=0.9,en-US;q=0.8,en;q=0.7",
    })

    r = session_get(s, LOGIN_URL)
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "html.parser")

    token_el = soup.select_one("input[name=_token]")
    if not token_el:
        raise RuntimeError("登入頁找不到 _token")

    token = token_el.get("value", "")

    resp = session_post(
        s,
        LOGIN_URL,
        data={
            "_token": token,
            "email": email,
            "password": password,
        },
        allow_redirects=True,
    )
    resp.raise_for_status()

    check = session_get(s, PURCHASE_URL, allow_redirects=True)
    check.raise_for_status()

    if "/login" in check.url:
        raise RuntimeError("登入失敗，請確認帳密")

    return s


# ========================
# List parsing
# ========================
def parse_purchase_list_page(html: str) -> List[Dict]:
    soup = BeautifulSoup(html, "html.parser")
    data = []

    rows = soup.select("table tbody tr")
    if not rows:
        rows = soup.select("tr")

    for tr in rows:
        txt = tr.get_text("\n", strip=True)

        m = re.search(r"(LC\d+)", txt)
        if not m:
            continue

        order_no = m.group(1)

        date_str = ""
        date_obj = None
        date_match = re.search(r"(\d{4}-\d{2}-\d{2})", txt)
        if date_match:
            date_str = date_match.group(1).replace("-", "/")
            date_obj = parse_date(date_match.group(1))

        status = ""
        status_code = ""
        if "未處理" in txt:
            status = "未處理"
            status_code = "0"
        elif "已處理" in txt:
            status = "已處理"
            status_code = "1"
        elif "已完成" in txt:
            status = "已完成"
            status_code = "2"

        purchase_status_name = ""
        purchase_status = ""
        if "待付款" in txt:
            purchase_status_name = "待付款"
            purchase_status = "0"
        elif "已付款" in txt:
            purchase_status_name = "已付款"
            purchase_status = "1"
        elif "取消訂單" in txt:
            purchase_status_name = "取消訂單"
            purchase_status = "2"
        elif "已退款" in txt:
            purchase_status_name = "已退款"
            purchase_status = "3"

        address = extract_address_from_text_block(txt)
        name = extract_name_from_text_block(txt)

        phone = ""
        m_phone = re.search(r"(09\d{8})", txt)
        if m_phone:
            phone = m_phone.group(1)

        edit_link = tr.select_one('a[href*="/purchase/edit/"]')
        edit_url = build_absolute_url(edit_link["href"]) if edit_link else ""

        data.append({
            "order_no": order_no,
            "date_str": date_str,
            "date_obj": date_obj,
            "status": status,
            "status_code": status_code,
            "address": address,
            "phone": phone,
            "name": name,
            "edit_url": edit_url,
            "purchase_status_name": purchase_status_name,
            "purchase_status": purchase_status,
        })

    dedup = {}
    for item in data:
        dedup[item["order_no"]] = item
    return list(dedup.values())


def search_paid_orders_by_phone(session, phone) -> List[Dict]:
    r = session_get(
        session,
        PURCHASE_URL,
        params={
            "keyword": "",
            "name": "",
            "phone": phone,
            "orderNo": "",
            "date_s": "",
            "date_e": "",
            "clean_date_s": "",
            "clean_date_e": "",
            "paid_at_s": "",
            "paid_at_e": "",
            "refundDateS": "",
            "refundDateE": "",
            "buy": "",
            "area_id": "",
            "isCharge": "",
            "isRefund": "",
            "payway": "",
            "purchase_status": "1",
            "progress_status": "",
            "invoiceStatus": "",
            "otherFee": "",
            "orderBy": "",
        },
    )
    r.raise_for_status()
    return parse_purchase_list_page(r.text)


def search_by_conditions(session, date_s: str, limit: int) -> List[Dict]:
    r = session_get(
        session,
        PURCHASE_URL,
        params={
            "keyword": "",
            "name": "",
            "phone": "",
            "orderNo": "",
            "date_s": date_s.replace("/", "-"),
            "date_e": date_s.replace("/", "-"),
            "clean_date_s": "",
            "clean_date_e": "",
            "paid_at_s": "",
            "paid_at_e": "",
            "refundDateS": "",
            "refundDateE": "",
            "buy": "",
            "area_id": "",
            "isCharge": "",
            "isRefund": "",
            "payway": "",
            "purchase_status": "1",
            "progress_status": "0",
            "invoiceStatus": "",
            "otherFee": "",
            "orderBy": "",
        },
    )
    r.raise_for_status()

    items = parse_purchase_list_page(r.text)
    filtered = [x for x in items if x.get("purchase_status") == "1" and x.get("status_code") == "0"]

    if limit and limit > 0:
        filtered = filtered[:limit]

    return filtered


# ========================
# Edit page parsing
# ========================
def parse_select_value(select_el):
    name = select_el.get("name", "")
    text = select_el.get_text(" ", strip=True)

    selected = select_el.select_one("option[selected]")
    if selected is not None:
        return selected.get("value", "")

    if name == "progress":
        if "已完成" in text:
            return "2"
        if "已處理" in text:
            return "1"
        if "未處理" in text:
            return "0"

    if name == "purchase_status":
        if "已退款" in text:
            return "3"
        if "取消訂單" in text:
            return "2"
        if "已付款" in text:
            return "1"
        if "待付款" in text:
            return "0"

    return ""


def parse_edit_page(session, edit_url, phone=""):
    params = {}
    if phone:
        params["phone"] = phone

    r = session_get(session, edit_url, params=params)
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "html.parser")

    form = soup.select_one("form")
    if not form:
        raise RuntimeError(f"找不到表單: {edit_url}")

    action = build_absolute_url(form.get("action") or edit_url)
    current_query_params = parse_query_params_from_url(r.url)

    fields = {}
    for el in form.select("input, textarea, select"):
        name = el.get("name")
        if not name:
            continue

        tag = el.name.lower()
        if tag == "textarea":
            fields[name] = el.text or ""
        elif tag == "select":
            fields[name] = parse_select_value(el)
        else:
            input_type = (el.get("type") or "text").lower()
            if input_type in ("checkbox", "radio"):
                if el.has_attr("checked"):
                    fields[name] = el.get("value", "on")
            else:
                fields[name] = el.get("value", "")

    if "_token" not in fields:
        token_el = soup.select_one("input[name=_token]")
        if token_el:
            fields["_token"] = token_el.get("value", "")

    notice_el = soup.select_one('textarea[name="notice"]')
    notice = notice_el.text.strip() if notice_el else str(fields.get("notice", "")).strip()

    page_text = soup.get_text("\n", strip=True)

    order_no = ""
    m = re.search(r"(LC\d+)", page_text)
    if m:
        order_no = m.group(1)

    customer_name = str(fields.get("name", "") or fields.get("customer_name", "") or "").strip()
    phone_value = str(fields.get("phone", "")).strip()
    address_value = str(fields.get("address", "")).strip()
    progress = str(fields.get("progress", "")).strip()
    purchase_status = str(fields.get("purchase_status", "")).strip()

    return {
        "action": action,
        "fields": fields,
        "notice": notice,
        "progress": progress,
        "purchase_status": purchase_status,
        "order_no": order_no,
        "edit_url": r.url,
        "query_params": current_query_params,
        "customer_name": customer_name,
        "phone": phone_value,
        "address": address_value,
        "page_text": page_text,
    }


# ========================
# Previous / targets
# ========================
def find_previous_processed(current_order_no, current_address, current_phone, current_service_date, items):
    matched = []
    for x in items:
        if x["order_no"] == current_order_no:
            continue
        if x.get("purchase_status") != "1":
            continue
        if x.get("status_code") not in ("1", "2"):
            continue
        if not x["date_obj"]:
            continue
        if current_service_date and x["date_obj"] >= current_service_date:
            continue
        if not same_address(x["address"], current_address):
            continue
        if current_phone and x.get("phone") and normalize_phone(x["phone"]) != normalize_phone(current_phone):
            continue
        matched.append(x)

    if not matched:
        return None

    matched.sort(key=lambda k: k["date_obj"], reverse=True)
    return matched[0]


def find_all_unprocessed_same_address(current_address, current_phone, items):
    targets = []
    for x in items:
        if x.get("status_code") != "0":
            continue
        if not same_address(x.get("address", ""), current_address):
            continue
        if current_phone and x.get("phone") and normalize_phone(x["phone"]) != normalize_phone(current_phone):
            continue
        targets.append(x)

    targets.sort(key=lambda k: (k.get("date_obj") or datetime.max, k.get("order_no", "")))
    return targets


def find_phone_mode_targets(items):
    by_address = {}
    for item in items:
        key = normalize_address(item["address"])
        if not key:
            continue
        by_address.setdefault(key, []).append(item)

    groups = []
    for _, address_items in by_address.items():
        unprocessed = [x for x in address_items if x.get("status_code") == "0" and x["date_obj"]]
        processed = [x for x in address_items if x.get("purchase_status") == "1" and x.get("status_code") in ("1", "2") and x["date_obj"]]

        if not unprocessed or not processed:
            continue

        unprocessed.sort(key=lambda x: x["date_obj"])
        processed.sort(key=lambda x: x["date_obj"], reverse=True)

        for target in unprocessed:
            prev = next((p for p in processed if p["date_obj"] < target["date_obj"]), None)
            if prev:
                groups.append({"target": target, "prev": prev})

    return groups


# ========================
# Submit / verify
# ========================
def submit_update(session, form_info, phone, new_notice):
    action = form_info["action"]
    fields = dict(form_info["fields"])
    query_params = dict(form_info.get("query_params", {}))

    fields["notice"] = new_notice
    fields["progress"] = "1"

    if phone and "phone" not in query_params:
        query_params["phone"] = phone

    resp = session_post(
        session,
        action,
        params=query_params,
        files={k: (None, str(v)) for k, v in fields.items()},
        headers={
            "Referer": form_info["edit_url"],
            "User-Agent": "Mozilla/5.0",
        },
        allow_redirects=True,
    )
    resp.raise_for_status()
    return resp


def verify_update(session, edit_url, phone, expected_notice):
    form = parse_edit_page(session, edit_url, phone)

    actual_notice = str(form.get("notice", "")).strip()
    actual_progress = str(form.get("progress", "")).strip()
    page_text = str(form.get("page_text", ""))

    notice_ok = actual_notice == str(expected_notice or "").strip()
    progress_ok = (
        actual_progress == "1"
        or "已處理" in page_text
    )

    return notice_ok and progress_ok, form


# ========================
# Core single case
# ========================
def process_single_case(session, order, name, phone, addr, date, log):
    current_service_date = parse_date(date)

    log(f"訂單: {order}")
    log(f"客戶: {name}")
    log(f"電話: {phone}")
    log(f"地址: {addr}")
    log(f"日期: {date}")

    items = search_paid_orders_by_phone(session, phone)

    log("\n[列表頁候選]")
    for i in items:
        log(f"{i['order_no']} {i['date_str']} {i['purchase_status_name']} {i['status']} {i['address']}")

    prev = find_previous_processed(order, addr, phone, current_service_date, items)
    if not prev:
        log("❌ 沒有上一筆已付款且相同電話地址的訂單")
        return {
            "ok": False,
            "prev_date": "",
            "prev_order": "",
            "prev_notice": "",
            "updated_orders": 0,
            "error": "沒有上一筆已付款且相同電話地址的訂單",
        }

    targets = find_all_unprocessed_same_address(addr, phone, items)
    if not targets:
        log("❌ 沒有未處理目標單")
        return {
            "ok": False,
            "prev_date": prev["date_str"],
            "prev_order": prev["order_no"],
            "prev_notice": "",
            "updated_orders": 0,
            "error": "沒有未處理目標單",
        }

    log(f"\n[上一筆] {prev['order_no']} {prev['date_str']} {prev['purchase_status_name']} {prev['status']} {prev['address']}")
    log(f"[本次共更新] {len(targets)} 筆未處理訂單")

    prev_form = parse_edit_page(session, prev["edit_url"], phone)
    prev_notice = str(prev_form.get("notice", "")).strip()

    if not prev_notice:
        log("❌ 上一筆找不到客服備註")
        return {
            "ok": False,
            "prev_date": prev["date_str"],
            "prev_order": prev["order_no"],
            "prev_notice": "",
            "updated_orders": 0,
            "error": "上一筆找不到客服備註",
        }

    success_count = 0
    fail_list = []

    for t in targets:
        log(f"👉 寫入 {t['order_no']} {t['date_str']} {t['purchase_status_name']} {t['status']} {t['address']}")
        try:
            target_form = parse_edit_page(session, t["edit_url"], phone)
            submit_update(session, target_form, phone, prev_notice)
            time.sleep(SLEEP_SECONDS)

            ok, vf = verify_update(session, t["edit_url"], phone, prev_notice)
            if ok:
                log(f"✅ 驗證成功 {t['order_no']}")
                success_count += 1
            else:
                log(
                    f"❌ 驗證失敗 {t['order_no']} "
                    f"(progress={vf.get('progress','')}, "
                    f"notice_head={str(vf.get('notice',''))[:30]})"
                )
                fail_list.append(t["order_no"])
        except Exception as e:
            log(f"❌ 寫入失敗 {t['order_no']}：{e}")
            fail_list.append(t["order_no"])

    if not fail_list:
        log(f"✅ 成功：已回填 {success_count} 筆")
        return {
            "ok": True,
            "prev_date": prev["date_str"],
            "prev_order": prev["order_no"],
            "prev_notice": prev_notice,
            "updated_orders": success_count,
            "error": "",
        }

    log(f"❌ 失敗：成功 {success_count} 筆，失敗 {len(fail_list)} 筆：{', '.join(fail_list)}")
    return {
        "ok": False,
        "prev_date": prev["date_str"],
        "prev_order": prev["order_no"],
        "prev_notice": prev_notice,
        "updated_orders": success_count,
        "error": f"失敗目標：{', '.join(fail_list)}",
    }


# ========================
# Result template
# ========================
def blank_result():
    return {
        "processed": 0,
        "success": 0,
        "failed": 0,
        "skipped": 0,
        "updated_orders": 0,
        "errors": [],
    }


# ========================
# Sheet mode
# ========================
def main(row_spec="2", force=False, ui_logger=None):
    ws = get_ws()
    log_ws = get_log_ws()
    rows = with_retry(ws.get_all_values)

    row_nums = parse_row_spec(row_spec)
    log = make_logger(ui_logger)

    result = blank_result()
    updates = []
    updated_row_numbers = []

    session = login()

    for r in row_nums:
        CURRENT_ROW_LOGS.clear()

        try:
            if r - 1 >= len(rows):
                log(f"\n===== 第{r}列 =====")
                log("❌ 超出資料範圍")
                result["failed"] += 1
                result["errors"].append(f"第{r}列：超出資料範圍")
                continue

            row = rows[r - 1]

            order = safe_cell(row, 2)
            date = safe_cell(row, 8)
            name = safe_cell(row, 13)
            addr = safe_cell(row, 14)
            phone = safe_cell(row, 15)
            v_status = safe_cell(row, 22)

            log(f"\n===== 第{r}列 =====")

            if v_status and not force:
                log(f"訂單: {order}")
                log(f"客戶: {name}")
                log(f"電話: {phone}")
                log(f"地址: {addr}")
                log(f"日期: {date}")
                log("⏭ 已有狀態，略過")
                result["skipped"] += 1
                continue

            if not order or not addr or not phone:
                log("❌ 缺少訂單 / 地址 / 電話")
                result["failed"] += 1
                result["errors"].append(f"第{r}列：缺少訂單 / 地址 / 電話")
                continue

            if result["processed"] == 0:
                log("[登入] 已登入")

            single_result = process_single_case(
                session=session,
                order=order,
                name=name,
                phone=phone,
                addr=addr,
                date=date,
                log=log,
            )

            updates.append({
                "range": f"S{r}:X{r}",
                "values": [[
                    single_result["prev_date"],
                    single_result["prev_order"],
                    single_result["prev_notice"],
                    "成功" if single_result["ok"] else "失敗",
                    "\n".join(CURRENT_ROW_LOGS),
                    single_result["prev_notice"],
                ]],
            })
            updated_row_numbers.append(r)

            append_log_row(
                log_ws=log_ws,
                source_type="BY列號",
                source_value=str(r),
                phone=phone,
                name=name,
                address=addr,
                current_order=order,
                current_date=date,
                prev_order=single_result["prev_order"],
                prev_date=single_result["prev_date"],
                prev_notice=single_result["prev_notice"],
                updated_orders=single_result["updated_orders"],
                status="成功" if single_result["ok"] else "失敗",
                error_msg=single_result["error"],
                full_log="\n".join(CURRENT_ROW_LOGS),
            )

            result["processed"] += 1
            result["updated_orders"] += single_result["updated_orders"]

            if single_result["ok"]:
                result["success"] += 1
            else:
                result["failed"] += 1
                result["errors"].append(f"第{r}列：{single_result['error']}")

        except Exception as e:
            log(f"❌ 例外錯誤：{e}")
            append_log_row(
                log_ws=log_ws,
                source_type="BY列號",
                source_value=str(r),
                phone="",
                name="",
                address="",
                current_order="",
                current_date="",
                prev_order="",
                prev_date="",
                prev_notice="",
                updated_orders=0,
                status="失敗",
                error_msg=str(e),
                full_log="\n".join(CURRENT_ROW_LOGS),
            )
            result["failed"] += 1
            result["errors"].append(f"第{r}列：{e}")

    if updates:
        with_retry(ws.batch_update, updates, value_input_option="RAW")
        apply_sheet_presentation(ws, updated_row_numbers)

    return result


# ========================
# Phone mode
# ========================
def main_by_phone(phone, ui_logger=None):
    CURRENT_ROW_LOGS.clear()
    log = make_logger(ui_logger)
    log_ws = get_log_ws()

    result = blank_result()
    phone = normalize_phone(phone)

    if not phone:
        msg = "電話號碼不可空白"
        log(f"❌ {msg}")
        result["failed"] = 1
        result["errors"].append(msg)
        return result

    log("\n===== BY電話 =====")
    log(f"輸入電話: {phone}")

    try:
        log("[登入] 已登入")
        session = login()
        items = search_paid_orders_by_phone(session, phone)
    except Exception as e:
        msg = f"登入/搜尋失敗：{e}"
        log(f"❌ {msg}")
        result["failed"] = 1
        result["errors"].append(msg)
        return result

    if not items:
        msg = "查無已付款訂單"
        log(f"❌ {msg}")
        result["failed"] = 1
        result["errors"].append(msg)
        append_log_row(
            log_ws, "BY電話", phone, phone, "", "", "", "", "", "", "", 0, "失敗", msg, "\n".join(CURRENT_ROW_LOGS)
        )
        return result

    log("\n[列表頁候選]")
    for i in items:
        log(f"{i['order_no']} {i['date_str']} {i['purchase_status_name']} {i['status']} {i['address']}")

    groups = find_phone_mode_targets(items)
    if not groups:
        msg = "找不到可處理的『未處理單 + 同地址前次已付款已處理單』"
        log(f"❌ {msg}")
        result["failed"] = 1
        result["errors"].append(msg)
        append_log_row(
            log_ws, "BY電話", phone, phone, "", "", "", "", "", "", "", 0, "失敗", msg, "\n".join(CURRENT_ROW_LOGS)
        )
        return result

    for idx, group in enumerate(groups, start=1):
        target = group["target"]
        prev = group["prev"]

        log(f"\n===== 處理第 {idx} 組 =====")
        log(f"[目標單] {target['order_no']} {target['date_str']} {target['purchase_status_name']} {target['status']} {target['address']}")
        log(f"[上一筆] {prev['order_no']} {prev['date_str']} {prev['purchase_status_name']} {prev['status']} {prev['address']}")

        try:
            prev_form = parse_edit_page(session, prev["edit_url"], phone)
            prev_notice = str(prev_form.get("notice", "")).strip()
            target_name = target.get("name", "") or parse_edit_page(session, target["edit_url"], phone).get("customer_name", "")

            if not prev_notice:
                log("❌ 上一筆找不到客服備註")
                result["failed"] += 1
                append_log_row(
                    log_ws, "BY電話", phone, phone, target_name, target["address"], target["order_no"], target["date_str"],
                    prev["order_no"], prev["date_str"], "", 0, "失敗", "上一筆找不到客服備註", "\n".join(CURRENT_ROW_LOGS)
                )
                continue

            same_address_unprocessed = find_all_unprocessed_same_address(target["address"], phone, items)
            log(f"[本次共更新] {len(same_address_unprocessed)} 筆未處理訂單")

            group_ok = 0
            group_fail = []

            for t in same_address_unprocessed:
                log(f"👉 寫入 {t['order_no']} {t['date_str']} {t['purchase_status_name']} {t['status']} {t['address']}")
                tf = parse_edit_page(session, t["edit_url"], phone)
                submit_update(session, tf, phone, prev_notice)
                time.sleep(SLEEP_SECONDS)

                ok, verified_form = verify_update(session, t["edit_url"], phone, prev_notice)
                if ok:
                    log(f"✅ 驗證成功 {t['order_no']}")
                    group_ok += 1
                else:
                    log(
                        f"❌ 驗證失敗 {t['order_no']} "
                        f"(progress={verified_form.get('progress','')}, "
                        f"notice_head={str(verified_form.get('notice',''))[:30]})"
                    )
                    group_fail.append(t["order_no"])

            result["processed"] += 1
            result["updated_orders"] += group_ok

            if not group_fail:
                log(f"✅ 成功：已回填 {group_ok} 筆")
                result["success"] += 1
                append_log_row(
                    log_ws, "BY電話", phone, phone, target_name, target["address"],
                    target["order_no"], target["date_str"], prev["order_no"], prev["date_str"], prev_notice,
                    group_ok, "成功", "", "\n".join(CURRENT_ROW_LOGS)
                )
            else:
                log(f"❌ 失敗：成功 {group_ok} 筆，失敗 {len(group_fail)} 筆：{', '.join(group_fail)}")
                result["failed"] += 1
                append_log_row(
                    log_ws, "BY電話", phone, phone, target_name, target["address"],
                    target["order_no"], target["date_str"], prev["order_no"], prev["date_str"], prev_notice,
                    group_ok, "失敗", f"失敗目標：{', '.join(group_fail)}", "\n".join(CURRENT_ROW_LOGS)
                )

        except Exception as e:
            log(f"❌ 處理失敗 {target['order_no']}：{e}")
            result["processed"] += 1
            result["failed"] += 1
            append_log_row(
                log_ws, "BY電話", phone, phone, target.get("name", ""), target["address"], target["order_no"], target["date_str"],
                prev["order_no"], prev["date_str"], "", 0, "失敗", str(e), "\n".join(CURRENT_ROW_LOGS)
            )

    if result["failed"] > 0:
        result["errors"].append(f"成功 {result['success']} 組，失敗 {result['failed']} 組")

    return result


# ========================
# Conditions mode
# ========================
def main_by_conditions(date_s: str, limit: int, ui_logger=None):
    CURRENT_ROW_LOGS.clear()
    log = make_logger(ui_logger)
    log_ws = get_log_ws()

    result = blank_result()

    if not parse_date(date_s):
        msg = "日期格式錯誤，請輸入 YYYY/MM/DD"
        log(f"❌ {msg}")
        result["failed"] = 1
        result["errors"].append(msg)
        return result

    log("\n===== BY搜尋條件 =====")
    log(f"訂購日期: {date_s}")
    log(f"條件: 付款狀態=已付款 + 服務狀態=未處理 + 每次只跑 {limit} 筆")

    try:
        log("[登入] 已登入")
        session = login()
        all_targets = search_by_conditions(session, date_s, limit=limit)
    except Exception as e:
        msg = f"搜尋失敗：{e}"
        log(f"❌ {msg}")
        result["failed"] = 1
        result["errors"].append(msg)
        return result

    if not all_targets:
        msg = "查無符合條件訂單"
        log(f"❌ {msg}")
        result["failed"] = 1
        result["errors"].append(msg)
        append_log_row(
            log_ws, "BY搜尋條件", date_s, "", "", "", "", "", "", "", "", 0, "失敗", msg, "\n".join(CURRENT_ROW_LOGS)
        )
        return result

    log("\n[搜尋結果]")
    for i in all_targets:
        log(f"{i['order_no']} {i['date_str']} {i['purchase_status_name']} {i['status']} {i['address']}")

    for idx, target in enumerate(all_targets, start=1):
        try:
            CURRENT_ROW_LOGS.append("")
            log(f"\n===== 處理第 {idx} 筆 =====")
            log(f"[目標單] {target['order_no']} {target['date_str']} {target['purchase_status_name']} {target['status']} {target['address']}")

            target_form = parse_edit_page(session, target["edit_url"], target.get("phone", ""))

            target_phone = normalize_phone(target_form.get("phone", "") or target.get("phone", ""))
            target_name = target_form.get("customer_name", "") or target.get("name", "")
            target_addr = target_form.get("address", "") or target["address"]

            log(f"訂單: {target['order_no']}")
            log(f"客戶: {target_name}")
            log(f"電話: {target_phone}")
            log(f"地址: {target_addr}")
            log(f"日期: {target['date_str']}")

            if not target_phone or not target_addr:
                raise RuntimeError("目標單缺少電話或地址")

            items = search_paid_orders_by_phone(session, target_phone)

            log("\n[列表頁候選]")
            for i in items:
                log(f"{i['order_no']} {i['date_str']} {i['purchase_status_name']} {i['status']} {i['address']}")

            prev = find_previous_processed(
                target["order_no"],
                target_addr,
                target_phone,
                target["date_obj"],
                items,
            )
            if not prev:
                raise RuntimeError("沒有上一筆已付款且相同電話地址的訂單")

            log(f"\n[上一筆] {prev['order_no']} {prev['date_str']} {prev['purchase_status_name']} {prev['status']} {prev['address']}")

            prev_form = parse_edit_page(session, prev["edit_url"], target_phone)
            prev_notice = str(prev_form.get("notice", "")).strip()
            if not prev_notice:
                raise RuntimeError("上一筆找不到客服備註")

            same_address_unprocessed = find_all_unprocessed_same_address(target_addr, target_phone, items)
            log(f"[本次共更新] {len(same_address_unprocessed)} 筆未處理訂單")

            group_ok = 0
            group_fail = []

            for t in same_address_unprocessed:
                log(f"👉 寫入 {t['order_no']} {t['date_str']} {t['purchase_status_name']} {t['status']} {t['address']}")
                tf = parse_edit_page(session, t["edit_url"], target_phone)
                submit_update(session, tf, target_phone, prev_notice)
                time.sleep(SLEEP_SECONDS)

                ok, verified_form = verify_update(session, t["edit_url"], target_phone, prev_notice)
                if ok:
                    log(f"✅ 驗證成功 {t['order_no']}")
                    group_ok += 1
                else:
                    log(
                        f"❌ 驗證失敗 {t['order_no']} "
                        f"(progress={verified_form.get('progress','')}, "
                        f"notice_head={str(verified_form.get('notice',''))[:30]})"
                    )
                    group_fail.append(t["order_no"])

            result["processed"] += 1
            result["updated_orders"] += group_ok

            if group_fail:
                log(f"❌ 失敗：成功 {group_ok} 筆，失敗 {len(group_fail)} 筆：{', '.join(group_fail)}")
                result["failed"] += 1
                append_log_row(
                    log_ws=log_ws,
                    source_type="BY搜尋條件",
                    source_value=date_s,
                    phone=target_phone,
                    name=target_name,
                    address=target_addr,
                    current_order=target["order_no"],
                    current_date=target["date_str"],
                    prev_order=prev["order_no"],
                    prev_date=prev["date_str"],
                    prev_notice=prev_notice,
                    updated_orders=group_ok,
                    status="失敗",
                    error_msg=f"失敗目標：{', '.join(group_fail)}",
                    full_log="\n".join(CURRENT_ROW_LOGS),
                )
            else:
                log(f"✅ 成功：已回填 {group_ok} 筆")
                result["success"] += 1
                append_log_row(
                    log_ws=log_ws,
                    source_type="BY搜尋條件",
                    source_value=date_s,
                    phone=target_phone,
                    name=target_name,
                    address=target_addr,
                    current_order=target["order_no"],
                    current_date=target["date_str"],
                    prev_order=prev["order_no"],
                    prev_date=prev["date_str"],
                    prev_notice=prev_notice,
                    updated_orders=group_ok,
                    status="成功",
                    error_msg="",
                    full_log="\n".join(CURRENT_ROW_LOGS),
                )

        except Exception as e:
            log(f"❌ 處理失敗 {target.get('order_no','')}：{e}")
            result["processed"] += 1
            result["failed"] += 1
            append_log_row(
                log_ws=log_ws,
                source_type="BY搜尋條件",
                source_value=date_s,
                phone=target.get("phone", ""),
                name=target.get("name", ""),
                address=target.get("address", ""),
                current_order=target.get("order_no", ""),
                current_date=target.get("date_str", ""),
                prev_order="",
                prev_date="",
                prev_notice="",
                updated_orders=0,
                status="失敗",
                error_msg=str(e),
                full_log="\n".join(CURRENT_ROW_LOGS),
            )

    if result["failed"] > 0:
        result["errors"].append(f"成功 {result['success']} 筆，失敗 {result['failed']} 筆")

    return result
