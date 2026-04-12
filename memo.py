# -*- coding: utf-8 -*-
import re
import time
from datetime import datetime
from typing import Optional, List, Dict, Callable

import requests
from bs4 import BeautifulSoup
import gspread
from google.oauth2.service_account import Credentials

import env

try:
    import streamlit as st
except Exception:
    st = None


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

RUNTIME_EMAIL = ""
RUNTIME_PASSWORD = ""


def set_runtime_credentials(email: str, password: str):
    global RUNTIME_EMAIL, RUNTIME_PASSWORD
    RUNTIME_EMAIL = (email or "").strip()
    RUNTIME_PASSWORD = (password or "").strip()


WORKSHEET_NAME = getattr(env, "WORKSHEET_NAME", "memo")
LOG_SHEET_NAME = getattr(env, "LOG_SHEET_NAME", "memo_log")
GOOGLE_SERVICE_ACCOUNT_FILE = getattr(env, "GOOGLE_SERVICE_ACCOUNT_FILE", "")
SLEEP_SECONDS = float(getattr(env, "SLEEP_SECONDS", 0.5))

REQUEST_TIMEOUT = 30
MAX_RETRIES = 3
RETRY_BACKOFF = 1.2

CURRENT_ROW_LOGS: List[str] = []


def make_logger(ui_logger: Optional[Callable[[str], None]] = None):
    def _log(msg: str):
        msg = str(msg)
        print(msg, flush=True)
        CURRENT_ROW_LOGS.append(msg)
        if ui_logger:
            ui_logger(msg)
    return _log


def blank_result():
    return {
        "processed": 0,
        "success": 0,
        "failed": 0,
        "skipped": 0,
        "updated_orders": 0,
        "errors": [],
    }


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


def parse_phone_list(text: str) -> List[str]:
    raw = re.split(r"[,\n;、，]+", str(text or ""))
    phones = []
    for x in raw:
        p = normalize_phone(x)
        if p:
            phones.append(p)
    return list(dict.fromkeys(phones))


def normalize_text(t: str) -> str:
    return re.sub(r"\s+", "", str(t or ""))


def normalize_address(addr: str) -> str:
    s = str(addr or "").strip()
    s = normalize_text(s)
    s = s.replace("臺", "台")
    s = s.replace("，", ",")
    s = s.replace("（", "(").replace("）", ")")
    s = s.replace("－", "-").replace("–", "-").replace("—", "-")
    s = s.replace("之", "-")
    s = s.replace("號-", "號")
    s = s.replace("樓-", "樓")
    s = s.replace(",", "")
    s = s.replace("　", "")
    return s


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


def extract_name_from_text_block(text: str) -> str:
    lines = [x.strip() for x in str(text or "").splitlines() if x.strip()]
    for line in lines:
        if re.search(r"^[\u4e00-\u9fff]{2,4}$", line):
            return line
    return ""


def extract_service_date_from_page_text(page_text: str) -> str:
    text = str(page_text or "")
    m = re.search(r"(\d{4}[/-]\d{2}[/-]\d{2})\s*\([一二三四五六日]\)", text)
    if m:
        return m.group(1).replace("-", "/")
    m = re.search(r"(\d{4}[/-]\d{2}[/-]\d{2})", text)
    if m:
        return m.group(1).replace("-", "/")
    return ""


def extract_address_from_text_block(text: str) -> str:
    text = str(text or "")
    city_pattern = (
        r"(?:台北市|臺北市|新北市|桃園市|新竹市|新竹縣|台中市|臺中市|"
        r"彰化縣|南投縣|雲林縣|嘉義市|嘉義縣|台南市|臺南市|高雄市|"
        r"屏東縣|宜蘭縣|花蓮縣|台東縣|臺東縣|基隆市)"
    )
    patterns = [
        rf"({city_pattern}[^\n]*?號(?:之\d+)?(?:\d+樓)?(?:之\d+)?(?:\d+室)?)",
        rf"({city_pattern}[^\n]*?樓之\d+)",
        rf"({city_pattern}[^\n]*?\d+樓)",
        rf"({city_pattern}[^\n]*?號)",
    ]
    for p in patterns:
        m = re.search(p, text)
        if m:
            return m.group(1).strip()
    return ""


def get_purchase_id_from_edit_url(edit_url: str) -> str:
    m = re.search(r"/purchase/edit/(\d+)", edit_url or "")
    return m.group(1) if m else ""


def display_service_date(item: Dict) -> str:
    return item.get("service_date") or item.get("raw_date_str") or ""


def item_service_date_obj(item: Dict):
    return item.get("service_date_obj") or item.get("raw_date_obj")


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
            "目前服務日期",
            "前次訂單",
            "前次服務日期",
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
    current_service_date: str,
    prev_order: str,
    prev_service_date: str,
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
            current_service_date,
            prev_order,
            prev_service_date,
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
                "startColumnIndex": 22,
                "endColumnIndex": 24,
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


def login(ui_logger=None):
    log = make_logger(ui_logger)

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

    log("[登入] 已登入")
    return s


def parse_purchase_row_text(txt: str) -> Dict:
    txt = str(txt or "")

    order_no = ""
    order_m = re.search(r"(LC\d+)", txt)
    if order_m:
        order_no = order_m.group(1)

    dates = re.findall(r"\d{4}-\d{2}-\d{2}", txt)
    raw_date_str = dates[0].replace("-", "/") if dates else ""
    raw_date_obj = parse_date(dates[0]) if dates else None

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
    if "待付款" in txt or "未付款" in txt:
        purchase_status_name = "未付款"
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

    return {
        "order_no": order_no,
        "raw_date_str": raw_date_str,
        "raw_date_obj": raw_date_obj,
        "status": status,
        "status_code": status_code,
        "address": address,
        "phone": phone,
        "name": name,
        "purchase_status_name": purchase_status_name,
        "purchase_status": purchase_status,
        "service_date": "",
        "service_date_obj": None,
    }


def parse_purchase_list_page(html: str) -> List[Dict]:
    soup = BeautifulSoup(html, "html.parser")
    data = []

    rows = soup.select("table tbody tr")
    if not rows:
        rows = soup.select("tr")

    for tr in rows:
        txt = tr.get_text("\n", strip=True)
        row_data = parse_purchase_row_text(txt)
        if not row_data["order_no"]:
            continue

        edit_link = tr.select_one('a[href*="/purchase/edit/"]')
        edit_url = f"{BASE_URL}{edit_link['href']}" if edit_link and edit_link.get("href", "").startswith("/") else (edit_link["href"] if edit_link else "")
        row_data["edit_url"] = edit_url
        row_data["purchase_id"] = get_purchase_id_from_edit_url(edit_url)
        data.append(row_data)

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


def search_all_orders_by_phone(session, phone) -> List[Dict]:
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
            "purchase_status": "",
            "progress_status": "",
            "invoiceStatus": "",
            "otherFee": "",
            "orderBy": "",
        },
    )
    r.raise_for_status()
    return parse_purchase_list_page(r.text)


def search_orders_by_order_no(session, order_no: str) -> List[Dict]:
    r = session_get(
        session,
        PURCHASE_URL,
        params={
            "keyword": "",
            "name": "",
            "phone": "",
            "orderNo": order_no,
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
            "purchase_status": "",
            "progress_status": "",
            "invoiceStatus": "",
            "otherFee": "",
            "orderBy": "",
        },
    )
    r.raise_for_status()
    return parse_purchase_list_page(r.text)


def search_by_conditions_once(session, date_mode: str, date_start: str, date_end: str, purchase_status_name: str, limit: int) -> List[Dict]:
    purchase_status_map = {
        "未付款": "0",
        "已付款": "1",
    }
    purchase_status = purchase_status_map.get(purchase_status_name, "")

    params = {
        "keyword": "",
        "name": "",
        "phone": "",
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
        "purchase_status": purchase_status,
        "progress_status": "0",
        "invoiceStatus": "",
        "otherFee": "",
        "orderBy": "",
    }

    if date_mode == "服務日期":
        params["clean_date_s"] = (date_start or "").replace("/", "-")
        params["clean_date_e"] = (date_end or "").replace("/", "-")
    else:
        params["date_s"] = (date_start or "").replace("/", "-")
        params["date_e"] = (date_end or "").replace("/", "-")

    r = session_get(session, PURCHASE_URL, params=params)
    r.raise_for_status()

    items = parse_purchase_list_page(r.text)
    filtered = [
        x for x in items
        if x.get("status_code") == "0" and (not purchase_status or x.get("purchase_status") == purchase_status)
    ]
    if limit and limit > 0:
        filtered = filtered[:limit]
    return filtered


def search_by_conditions(session, date_mode: str, date_start: str, date_end: str, purchase_status_name: str, limit: int) -> List[Dict]:
    if purchase_status_name == "全部":
        unpaid = search_by_conditions_once(session, date_mode, date_start, date_end, "未付款", limit)
        paid = search_by_conditions_once(session, date_mode, date_start, date_end, "已付款", limit)
        merged = {}
        for x in unpaid + paid:
            merged[x["order_no"]] = x
        items = list(merged.values())
        items.sort(key=lambda x: (display_service_date(x), x.get("order_no", "")))
        if limit and limit > 0:
            items = items[:limit]
        return items
    return search_by_conditions_once(session, date_mode, date_start, date_end, purchase_status_name, limit)


def parse_select_value(select_el):
    name = select_el.get("name", "")
    text = select_el.get_text(" ", strip=True)

    selected = select_el.select_one("option[selected]")
    if selected is not None:
        return selected.get("value", "")

    if name in ("progress", "progress_status"):
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
        if "待付款" in text or "未付款" in text:
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

    action = form.get("action") or edit_url
    if action.startswith("/"):
        action = f"{BASE_URL}{action}"

    current_query_params = {}
    if "?" in r.url:
        query = r.url.split("?", 1)[1]
        for pair in query.split("&"):
            if "=" in pair:
                k, v = pair.split("=", 1)
                current_query_params[k] = v

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

    page_text = soup.get_text("\n", strip=True)

    notice_el = soup.select_one('textarea[name="notice"]')
    notice = notice_el.text.strip() if notice_el else str(fields.get("notice", "")).strip()

    order_no = ""
    m = re.search(r"(LC\d+)", page_text)
    if m:
        order_no = m.group(1)

    service_date = extract_service_date_from_page_text(page_text)

    customer_name = str(fields.get("name", "") or fields.get("customer_name", "") or "").strip()
    phone_value = str(fields.get("phone", "")).strip()
    address_value = str(fields.get("address", "")).strip()
    progress = str(fields.get("progress", "") or fields.get("progress_status", "")).strip()
    purchase_status = str(fields.get("purchase_status", "")).strip()

    purchase_id = get_purchase_id_from_edit_url(edit_url)

    purchase_status_name = ""
    if purchase_status == "0":
        purchase_status_name = "未付款"
    elif purchase_status == "1":
        purchase_status_name = "已付款"
    elif purchase_status == "2":
        purchase_status_name = "取消訂單"
    elif purchase_status == "3":
        purchase_status_name = "已退款"

    return {
        "action": action,
        "fields": fields,
        "notice": notice,
        "progress": progress,
        "purchase_status": purchase_status,
        "purchase_status_name": purchase_status_name,
        "order_no": order_no,
        "purchase_id": purchase_id,
        "edit_url": r.url,
        "query_params": current_query_params,
        "customer_name": customer_name,
        "phone": phone_value,
        "address": address_value,
        "page_text": page_text,
        "service_date": service_date,
        "service_date_obj": parse_date(service_date),
    }


def enrich_item_from_detail(session, item: Dict, phone="") -> Dict:
    if not item.get("edit_url"):
        return item

    detail = parse_edit_page(session, item["edit_url"], phone)
    item["service_date"] = detail.get("service_date", "") or item.get("service_date", "")
    item["service_date_obj"] = detail.get("service_date_obj") or item.get("service_date_obj")
    item["phone"] = detail.get("phone", "") or item.get("phone", "")
    item["name"] = detail.get("customer_name", "") or item.get("name", "")
    item["address"] = detail.get("address", "") or item.get("address", "")
    item["purchase_status_name"] = detail.get("purchase_status_name", "") or item.get("purchase_status_name", "")
    return item


def enrich_items_from_detail(session, items: List[Dict], phone="") -> List[Dict]:
    result = []
    for item in items:
        try:
            result.append(enrich_item_from_detail(session, item, phone))
        except Exception:
            result.append(item)
    return result


def dedupe_items_by_address(items: List[Dict]) -> List[Dict]:
    by_address = {}

    for item in items:
        key = normalize_address(item.get("address", ""))
        if not key:
            continue

        existed = by_address.get(key)
        current_dt = item_service_date_obj(item) or datetime.min

        if not existed:
            by_address[key] = item
            continue

        existed_dt = item_service_date_obj(existed) or datetime.min
        if current_dt > existed_dt:
            by_address[key] = item

    result = list(by_address.values())
    result.sort(key=lambda x: item_service_date_obj(x) or datetime.min, reverse=True)
    return result


def log_candidate_addresses(log, items: List[Dict]):
    log("\n[列表頁候選 - 各地址一筆]")
    deduped = dedupe_items_by_address(items)
    for i in deduped:
        log(
            f"{display_service_date(i)} {i['order_no']} "
            f"{i.get('name','')} {i.get('purchase_status_name','')} {i['status']} {i['address']}"
        )


def find_previous_processed_verbose(current_order_no, current_address, current_phone, current_service_date, items):
    same_phone = []
    same_address_items = []
    earlier_items = []
    paid_items = []
    processed_items = []

    for x in items:
        if x["order_no"] == current_order_no:
            continue

        if current_phone and x.get("phone") and normalize_phone(x["phone"]) != normalize_phone(current_phone):
            continue
        same_phone.append(x)

        if not same_address(x.get("address", ""), current_address):
            continue
        same_address_items.append(x)

        dt = item_service_date_obj(x)
        if not dt:
            continue

        if current_service_date and dt >= current_service_date:
            continue
        earlier_items.append(x)

        if x.get("purchase_status") != "1":
            continue
        paid_items.append(x)

        if x.get("status_code") not in ("1", "2"):
            continue
        processed_items.append(x)

    if processed_items:
        processed_items.sort(key=lambda k: item_service_date_obj(k), reverse=True)
        return processed_items[0], ""

    if not same_phone:
        return None, "沒有相同電話的歷史訂單"
    if not same_address_items:
        return None, "有相同電話，但沒有相同地址的歷史訂單"
    if not earlier_items:
        return None, "有相同電話地址的訂單，但沒有更早的服務日期"
    if not paid_items:
        return None, "有相同電話地址且更早的訂單，但付款狀態不是已付款"

    return None, "有相同電話地址且已付款的更早訂單，但服務狀態不是已處理"


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

    targets.sort(key=lambda k: (item_service_date_obj(k) or datetime.max, k.get("order_no", "")))
    return targets


def submit_update(session, form_info, phone, new_notice):
    action = form_info["action"]
    fields = dict(form_info["fields"])
    query_params = dict(form_info.get("query_params", {}))

    if "notice" in fields:
        fields["notice"] = new_notice
    elif "purchase[notice]" in fields:
        fields["purchase[notice]"] = new_notice
    else:
        fields["notice"] = new_notice

    updated_progress = False
    for key in ["progress", "progress_status", "purchase[progress]", "purchase[progress_status]"]:
        if key in fields:
            fields[key] = "1"
            updated_progress = True
    if not updated_progress:
        fields["progress"] = "1"

    if form_info.get("purchase_id") and "id" not in query_params:
        query_params["id"] = form_info["purchase_id"]

    if phone and "phone" not in query_params:
        query_params["phone"] = phone

    resp = session_post(
        session,
        action,
        params=query_params,
        data=fields,
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

    norm_expected = normalize_text(expected_notice)
    norm_actual = normalize_text(actual_notice)

    notice_ok = norm_actual == norm_expected or (norm_expected[:20] and norm_expected[:20] in norm_actual)
    progress_ok = actual_progress == "1" or "已處理" in page_text

    return notice_ok and progress_ok, form


def process_single_case(session, order, name, phone, addr, date, log):
    current_service_date = parse_date(date)

    log(f"訂單: {order}")
    log(f"客戶: {name}")
    log(f"電話: {phone}")
    log(f"地址: {addr}")
    log(f"日期: {date}")

    items = search_paid_orders_by_phone(session, phone)
    items = enrich_items_from_detail(session, items, phone)

    log_candidate_addresses(log, items)

    prev, reason = find_previous_processed_verbose(order, addr, phone, current_service_date, items)
    if not prev:
        reason_text = f"❌ 處理失敗 {order}：{reason}"
        log(reason_text)
        return {
            "ok": False,
            "prev_date": "",
            "prev_order": "",
            "prev_notice": "",
            "updated_orders": 0,
            "error": reason_text,
        }

    targets = find_all_unprocessed_same_address(addr, phone, items)
    if not targets:
        reason_text = f"❌ 處理失敗 {order}：沒有未處理目標單"
        log(reason_text)
        return {
            "ok": False,
            "prev_date": display_service_date(prev),
            "prev_order": prev["order_no"],
            "prev_notice": "",
            "updated_orders": 0,
            "error": reason_text,
        }

    log(f"\n[上一筆] {display_service_date(prev)} {prev['order_no']} {prev.get('name','')} {prev['purchase_status_name']} {prev['status']} {prev['address']}")
    log(f"[本次共更新] {len(targets)} 筆未處理訂單")

    prev_form = parse_edit_page(session, prev["edit_url"], phone)
    prev_notice = str(prev_form.get("notice", "")).strip()

    if not prev_notice:
        reason_text = f"❌ 處理失敗 {order}：上一筆找不到客服備註"
        log(reason_text)
        return {
            "ok": False,
            "prev_date": display_service_date(prev),
            "prev_order": prev["order_no"],
            "prev_notice": "",
            "updated_orders": 0,
            "error": reason_text,
        }

    success_count = 0
    fail_list = []

    for t in targets:
        log(f"👉 寫入 {display_service_date(t)} {t['order_no']} {t.get('name','')} {t['purchase_status_name']} {t['status']} {t['address']}")
        try:
            target_form = parse_edit_page(session, t["edit_url"], phone)
            submit_update(session, target_form, phone, prev_notice)
            time.sleep(SLEEP_SECONDS)

            ok, vf = verify_update(session, t["edit_url"], phone, prev_notice)
            if ok:
                log(f"✅ 驗證成功 {t['order_no']}")
                success_count += 1
            else:
                log(f"❌ 驗證失敗 {t['order_no']} (progress={vf.get('progress','')}, notice_head={str(vf.get('notice',''))[:30]})")
                fail_list.append(t["order_no"])
        except Exception as e:
            log(f"❌ 寫入失敗 {t['order_no']}：{e}")
            fail_list.append(t["order_no"])

    if not fail_list:
        log(f"✅ 成功：已回填 {success_count} 筆")
        return {
            "ok": True,
            "prev_date": display_service_date(prev),
            "prev_order": prev["order_no"],
            "prev_notice": prev_notice,
            "updated_orders": success_count,
            "error": "",
        }

    reason_text = f"❌ 處理失敗 {order}：失敗目標 {', '.join(fail_list)}"
    log(reason_text)
    return {
        "ok": False,
        "prev_date": display_service_date(prev),
        "prev_order": prev["order_no"],
        "prev_notice": prev_notice,
        "updated_orders": success_count,
        "error": reason_text,
    }


def build_preview_row(target: Dict, history_items: List[Dict]) -> Dict:
    addr = target.get("address", "")
    same_count = 0
    diff_count = 0

    for x in history_items:
        if x.get("order_no") == target.get("order_no"):
            continue
        if same_address(x.get("address", ""), addr):
            same_count += 1
        else:
            diff_count += 1

    return {
        "order_id": target.get("order_no", ""),
        "customer_name": target.get("name", ""),
        "phone": target.get("phone", ""),
        "address": addr,
        "service_date": display_service_date(target),
        "same_address_count": same_count,
        "different_address_count": diff_count,
        "has_same_address_history": same_count > 0,
        "purchase_status_name": target.get("purchase_status_name", ""),
    }


def preview_by_phone(phone, ui_logger=None):
    CURRENT_ROW_LOGS.clear()
    log = make_logger(ui_logger)

    phone = normalize_phone(phone)
    if not phone:
        raise RuntimeError("電話不可為空")

    log("\n===== 預覽 BY電話 =====")
    log(f"輸入電話: {phone}")

    session = login(ui_logger=ui_logger)
    items = search_all_orders_by_phone(session, phone)
    items = enrich_items_from_detail(session, items, phone)

    if not items:
        return []

    preview = []
    for item in items:
        addr = item.get("address", "")
        if not addr:
            continue
        preview.append(build_preview_row(item, items))

    dedup = {}
    for row in preview:
        dedup[row["order_id"]] = row

    result = list(dedup.values())
    result.sort(
        key=lambda x: (
            0 if x.get("has_same_address_history") else 1,
            0 if x.get("purchase_status_name") == "已付款" else 1,
            -(x.get("same_address_count", 0) or 0),
            str(x.get("service_date", "")),
            str(x.get("order_id", "")),
        )
    )
    return result


def preview_by_phone_multi(phone_text, ui_logger=None):
    phones = parse_phone_list(phone_text)
    if not phones:
        raise RuntimeError("請輸入至少一支有效電話")

    all_rows = []
    seen = set()
    for phone in phones:
        rows = preview_by_phone(phone, ui_logger=ui_logger)
        for row in rows:
            oid = row.get("order_id", "")
            if oid and oid not in seen:
                seen.add(oid)
                all_rows.append(row)

    all_rows.sort(
        key=lambda x: (
            0 if x.get("has_same_address_history") else 1,
            0 if x.get("purchase_status_name") == "已付款" else 1,
            -(x.get("same_address_count", 0) or 0),
            str(x.get("service_date", "")),
            str(x.get("order_id", "")),
        )
    )
    return all_rows


def preview_by_conditions(date_mode, date_start, date_end, purchase_status_name, limit, ui_logger=None):
    CURRENT_ROW_LOGS.clear()
    log = make_logger(ui_logger)

    if not date_start and not date_end:
        raise RuntimeError("請至少選擇開始日期或結束日期")

    log("\n===== 預覽 BY搜尋條件 =====")
    log(f"日期條件: {date_mode}")
    log(f"日期區間: {date_start or '(空)'} ~ {date_end or '(空)'}")
    log(f"付款狀態: {purchase_status_name}")
    log(f"處理筆數: {limit}")

    session = login(ui_logger=ui_logger)
    items = search_by_conditions(session, date_mode, date_start, date_end, purchase_status_name, limit)
    items = enrich_items_from_detail(session, items)

    if not items:
        return []

    phone_cache = {}
    preview = []

    for item in items:
        phone = normalize_phone(item.get("phone", ""))
        addr = item.get("address", "")

        if not phone or not addr:
            preview.append({
                "order_id": item.get("order_no", ""),
                "customer_name": item.get("name", ""),
                "phone": phone,
                "address": addr,
                "service_date": display_service_date(item),
                "same_address_count": 0,
                "different_address_count": 0,
                "has_same_address_history": False,
                "purchase_status_name": item.get("purchase_status_name", ""),
            })
            continue

        if phone not in phone_cache:
            history = search_all_orders_by_phone(session, phone)
            history = enrich_items_from_detail(session, history, phone)
            phone_cache[phone] = history

        preview.append(build_preview_row(item, phone_cache[phone]))

    dedup = {}
    for row in preview:
        dedup[row["order_id"]] = row

    result = list(dedup.values())
    result.sort(
        key=lambda x: (
            0 if x.get("has_same_address_history") else 1,
            0 if x.get("purchase_status_name") == "已付款" else 1,
            -(x.get("same_address_count", 0) or 0),
            str(x.get("service_date", "")),
            str(x.get("order_id", "")),
        )
    )
    return result


def execute_target_order(session, target: Dict, source_type: str, source_value: str, log, log_ws):
    CURRENT_ROW_LOGS.clear()

    target_order_no = target.get("order_no", "")
    target_form = parse_edit_page(session, target["edit_url"], target.get("phone", ""))

    target_phone = normalize_phone(target_form.get("phone", "") or target.get("phone", ""))
    target_name = target_form.get("customer_name", "") or target.get("name", "")
    target_addr = target_form.get("address", "") or target.get("address", "")
    target_service_date = target_form.get("service_date", "") or display_service_date(target)

    log(f"\n===== 處理訂單 {target_order_no} =====")
    log(f"訂單: {target_order_no}")
    log(f"客戶: {target_name}")
    log(f"電話: {target_phone}")
    log(f"地址: {target_addr}")
    log(f"日期: {target_service_date}")

    if not target_phone or not target_addr:
        raise RuntimeError(f"❌ 處理失敗 {target_order_no}：目標單缺少電話或地址")

    items = search_paid_orders_by_phone(session, target_phone)
    items = enrich_items_from_detail(session, items, target_phone)

    log_candidate_addresses(log, items)

    prev, reason = find_previous_processed_verbose(
        current_order_no=target_order_no,
        current_address=target_addr,
        current_phone=target_phone,
        current_service_date=target_form.get("service_date_obj"),
        items=items,
    )
    if not prev:
        raise RuntimeError(f"❌ 處理失敗 {target_order_no}：{reason}")

    log(f"\n[上一筆] {display_service_date(prev)} {prev['order_no']} {prev.get('name','')} {prev['purchase_status_name']} {prev['status']} {prev['address']}")

    prev_form = parse_edit_page(session, prev["edit_url"], target_phone)
    prev_notice = str(prev_form.get("notice", "")).strip()
    if not prev_notice:
        raise RuntimeError(f"❌ 處理失敗 {target_order_no}：上一筆找不到客服備註")

    same_address_unprocessed = find_all_unprocessed_same_address(target_addr, target_phone, items)
    log(f"[本次共更新] {len(same_address_unprocessed)} 筆未處理訂單")

    group_ok = 0
    group_fail = []

    for t in same_address_unprocessed:
        log(f"👉 寫入 {display_service_date(t)} {t['order_no']} {t.get('name','')} {t['purchase_status_name']} {t['status']} {t['address']}")
        tf = parse_edit_page(session, t["edit_url"], target_phone)
        submit_update(session, tf, target_phone, prev_notice)
        time.sleep(SLEEP_SECONDS)

        ok, verified_form = verify_update(session, t["edit_url"], target_phone, prev_notice)
        if ok:
            log(f"✅ 驗證成功 {t['order_no']}")
            group_ok += 1
        else:
            log(f"❌ 驗證失敗 {t['order_no']} (progress={verified_form.get('progress','')}, notice_head={str(verified_form.get('notice',''))[:30]})")
            group_fail.append(t["order_no"])

    if group_fail:
        reason_text = f"❌ 處理失敗 {target_order_no}：失敗目標 {', '.join(group_fail)}"
        log(reason_text)
        append_log_row(
            log_ws=log_ws,
            source_type=source_type,
            source_value=source_value,
            phone=target_phone,
            name=target_name,
            address=target_addr,
            current_order=target_order_no,
            current_service_date=target_service_date,
            prev_order=prev["order_no"],
            prev_service_date=display_service_date(prev),
            prev_notice=prev_notice,
            updated_orders=group_ok,
            status="失敗",
            error_msg=reason_text,
            full_log="\n".join(CURRENT_ROW_LOGS),
        )
        return {"ok": False, "error": reason_text, "updated_orders": group_ok}

    log(f"✅ 成功：已回填 {group_ok} 筆")
    append_log_row(
        log_ws=log_ws,
        source_type=source_type,
        source_value=source_value,
        phone=target_phone,
        name=target_name,
        address=target_addr,
        current_order=target_order_no,
        current_service_date=target_service_date,
        prev_order=prev["order_no"],
        prev_service_date=display_service_date(prev),
        prev_notice=prev_notice,
        updated_orders=group_ok,
        status="成功",
        error_msg="",
        full_log="\n".join(CURRENT_ROW_LOGS),
    )
    return {"ok": True, "error": "", "updated_orders": group_ok}


def get_sheet_summary(ui_logger=None):
    log = make_logger(ui_logger)
    ws = get_ws()
    rows = with_retry(ws.get_all_values)

    total_rows = max(0, len(rows) - 1)
    pending_rows = 0
    done_rows = 0

    for row in rows[1:]:
        v_status = safe_cell(row, 22)
        if v_status:
            done_rows += 1
        else:
            order = safe_cell(row, 2)
            addr = safe_cell(row, 14)
            phone = safe_cell(row, 15)
            if order and addr and phone:
                pending_rows += 1

    log(f"總筆數={total_rows}, 未處理={pending_rows}, 已處理={done_rows}")
    return {
        "total_rows": total_rows,
        "pending_rows": pending_rows,
        "done_rows": done_rows,
    }


def get_first_n_pending_rows(limit: int) -> List[int]:
    ws = get_ws()
    rows = with_retry(ws.get_all_values)

    row_nums = []
    for idx, row in enumerate(rows[1:], start=2):
        v_status = safe_cell(row, 22)
        order = safe_cell(row, 2)
        addr = safe_cell(row, 14)
        phone = safe_cell(row, 15)

        if not v_status and order and addr and phone:
            row_nums.append(idx)

        if len(row_nums) >= limit:
            break

    return row_nums


def main_first_n_pending(limit: int, ui_logger=None):
    row_nums = get_first_n_pending_rows(limit)
    if not row_nums:
        return {
            "processed": 0,
            "success": 0,
            "failed": 0,
            "skipped": 0,
            "updated_orders": 0,
            "errors": ["沒有可處理的未處理列"],
        }
    row_spec = ",".join(str(x) for x in row_nums)
    return main(row_spec=row_spec, force=False, ui_logger=ui_logger)


def main(row_spec="2", force=False, ui_logger=None):
    ws = get_ws()
    log_ws = get_log_ws()
    rows = with_retry(ws.get_all_values)

    row_nums = parse_row_spec(row_spec)
    log = make_logger(ui_logger)

    result = blank_result()
    updates = []
    updated_row_numbers = []

    session = login(ui_logger=ui_logger)

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
                reason_text = f"❌ 處理失敗 第{r}列：缺少訂單 / 地址 / 電話"
                log(reason_text)
                result["failed"] += 1
                result["errors"].append(reason_text)
                continue

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
                current_service_date=date,
                prev_order=single_result["prev_order"],
                prev_service_date=single_result["prev_date"],
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
                result["errors"].append(single_result["error"])

        except Exception as e:
            reason_text = f"❌ 處理失敗 第{r}列：{e}"
            log(reason_text)
            append_log_row(
                log_ws=log_ws,
                source_type="BY列號",
                source_value=str(r),
                phone="",
                name="",
                address="",
                current_order="",
                current_service_date="",
                prev_order="",
                prev_service_date="",
                prev_notice="",
                updated_orders=0,
                status="失敗",
                error_msg=reason_text,
                full_log="\n".join(CURRENT_ROW_LOGS),
            )
            result["failed"] += 1
            result["errors"].append(reason_text)

    if updates:
        with_retry(ws.batch_update, updates, value_input_option="RAW")
        apply_sheet_presentation(ws, updated_row_numbers)

    return result


def main_by_selected_order_ids(order_ids, ui_logger=None):
    CURRENT_ROW_LOGS.clear()
    log = make_logger(ui_logger)
    log_ws = get_log_ws()
    result = blank_result()

    if not order_ids:
        msg = "❌ 處理失敗：未提供任何訂單"
        log(msg)
        result["failed"] = 1
        result["errors"].append(msg)
        return result

    session = login(ui_logger=ui_logger)
    source_type = "BY勾選執行"
    source_value = ",".join(order_ids[:50])

    for order_id in order_ids:
        try:
            items = search_orders_by_order_no(session, order_id)
            items = enrich_items_from_detail(session, items)

            target = next((x for x in items if x.get("order_no") == order_id), None)
            if not target:
                raise RuntimeError(f"❌ 處理失敗 {order_id}：找不到訂單")

            single_result = execute_target_order(
                session=session,
                target=target,
                source_type=source_type,
                source_value=source_value,
                log=log,
                log_ws=log_ws,
            )

            result["processed"] += 1
            result["updated_orders"] += single_result["updated_orders"]

            if single_result["ok"]:
                result["success"] += 1
            else:
                result["failed"] += 1
                result["errors"].append(single_result["error"])

        except Exception as e:
            reason_text = str(e)
            log(reason_text)
            result["processed"] += 1
            result["failed"] += 1
            result["errors"].append(reason_text)

            append_log_row(
                log_ws=log_ws,
                source_type=source_type,
                source_value=source_value,
                phone="",
                name="",
                address="",
                current_order=order_id,
                current_service_date="",
                prev_order="",
                prev_service_date="",
                prev_notice="",
                updated_orders=0,
                status="失敗",
                error_msg=reason_text,
                full_log="\n".join(CURRENT_ROW_LOGS),
            )

    return result
