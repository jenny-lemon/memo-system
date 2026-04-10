# -*- coding: utf-8 -*-
import re
import time
from datetime import datetime
from urllib.parse import urljoin, urlparse, parse_qs
from typing import Optional, List, Dict, Callable

import gspread
import requests
from bs4 import BeautifulSoup
from google.oauth2.service_account import Credentials

import accounts
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
    ENV_NAME = env_name
    BASE_URL = BASE_URL_DEV if env_name == "dev" else BASE_URL_PROD
    BASE_URL = BASE_URL.rstrip("/")
    LOGIN_URL = f"{BASE_URL}/login"
    PURCHASE_URL = f"{BASE_URL}/purchase"


set_env(ENV_NAME)

# ========================
# Google
# ========================
WORKSHEET_NAME = getattr(env, "WORKSHEET_NAME", "memo")
LOG_SHEET_NAME = getattr(env, "LOG_SHEET_NAME", "memo_log")
SLEEP_SECONDS = getattr(env, "SLEEP_SECONDS", 0.5)
GOOGLE_SERVICE_ACCOUNT_FILE = getattr(env, "GOOGLE_SERVICE_ACCOUNT_FILE", "")

CURRENT_ROW_LOGS: List[str] = []


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
    sh = get_spreadsheet()
    return sh.worksheet(WORKSHEET_NAME)


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
            "狀態",
            "錯誤訊息",
            "完整LOG",
        ])
        return ws


# ========================
# logger
# ========================
def make_logger(ui_logger: Optional[Callable[[str], None]] = None):
    def _log(msg: str):
        print(msg, flush=True)
        CURRENT_ROW_LOGS.append(msg)
        if ui_logger:
            ui_logger(msg)
    return _log


# ========================
# 工具
# ========================
def normalize_phone(p):
    return re.sub(r"\D+", "", str(p or ""))


def normalize_text(t):
    return re.sub(r"\s+", "", str(t or ""))


def parse_date(t):
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


def parse_row_spec(spec):
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


def safe_cell(row, idx_1_based):
    i = idx_1_based - 1
    return str(row[i]).strip() if i < len(row) else ""


def same_address(a: str, b: str) -> bool:
    na = normalize_text(a)
    nb = normalize_text(b)
    if not na or not nb:
        return False
    return na == nb or na in nb or nb in na


def match_region_by_address(address: str) -> Optional[str]:
    addr = str(address or "")
    for region, cfg in accounts.ACCOUNTS.items():
        for kw in cfg.get("address_keywords", []):
            if kw and kw in addr:
                return region
    return None


def extract_address_from_text_block(text: str) -> str:
    lines = [x.strip() for x in str(text or "").splitlines() if x.strip()]
    for i, line in enumerate(lines):
        if any(k in line for k in ["市", "縣", "區", "鄉", "鎮", "路", "街", "段", "巷", "弄", "號"]):
            if "付款" in line or "服務狀態" in line or "付款狀態" in line:
                continue
            addr = line
            if i + 1 < len(lines):
                next_line = lines[i + 1]
                if any(k in next_line for k in ["樓", "室", "之", "A", "B", "C", "D"]):
                    addr += next_line
            return addr
    return ""


def parse_query_params_from_url(url: str) -> Dict[str, str]:
    parsed = urlparse(url)
    qs = parse_qs(parsed.query)
    result = {}
    for k, v in qs.items():
        if v:
            result[k] = v[0]
    return result


def clip_text(text: str, limit: int = 50000) -> str:
    text = str(text or "")
    return text[:limit]


# ========================
# 後台登入
# ========================
def login(region):
    cfg = accounts.ACCOUNTS[region]
    email = str(cfg.get("email", "")).strip()
    password = str(cfg.get("password", "")).strip()

    if not email or not password:
        raise RuntimeError(f"{region} 缺少 email/password")

    s = requests.Session()
    s.headers.update({
        "User-Agent": "Mozilla/5.0",
        "Accept-Language": "zh-TW,zh;q=0.9,en-US;q=0.8,en;q=0.7",
    })

    r = s.get(LOGIN_URL, timeout=30)
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "html.parser")

    token_el = soup.select_one("input[name=_token]")
    if not token_el:
        raise RuntimeError("登入頁找不到 _token")

    token = token_el.get("value", "")

    resp = s.post(
        LOGIN_URL,
        data={
            "_token": token,
            "email": email,
            "password": password,
        },
        timeout=30,
        allow_redirects=True,
    )
    resp.raise_for_status()

    check = s.get(PURCHASE_URL, timeout=30, allow_redirects=True)
    check.raise_for_status()

    if "/login" in check.url:
        raise RuntimeError(f"{region} 登入失敗")

    return s


# ========================
# 搜尋：手機 + 已付款
# ========================
def search_paid_orders_by_phone(session, phone) -> List[Dict]:
    r = session.get(
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
            "purchase_status": "1",   # 已付款
            "progress_status": "",
            "invoiceStatus": "",
            "otherFee": "",
            "orderBy": "",
        },
        timeout=30,
    )
    r.raise_for_status()
    return parse_purchase_list_page(r.text)


# ========================
# 搜尋：訂購日期 + 非取消 + 未處理
# ========================
def search_by_conditions(session, date_s: str, limit: int) -> List[Dict]:
    """
    條件：
    - 訂購日期 = date_s
    - 付款狀態非取消 => 這裡用 purchase_status 不放 2
      實務上後台列表通常不能直接表達 != 2，所以先查當日未處理，再在解析結果排除 status=取消訂單
    - 服務狀態未處理 => progress_status=0
    """
    r = session.get(
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
            "purchase_status": "",
            "progress_status": "0",   # 未處理
            "invoiceStatus": "",
            "otherFee": "",
            "orderBy": "",
        },
        timeout=30,
    )
    r.raise_for_status()

    items = parse_purchase_list_page(r.text)

    filtered = []
    for item in items:
        # 排除取消
        if item.get("purchase_status_name") == "取消訂單":
            continue
        if item.get("status") != "未處理":
            continue
        filtered.append(item)

    if limit and limit > 0:
        filtered = filtered[:limit]

    return filtered


# ========================
# 共用：解析列表頁
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
        if "未處理" in txt:
            status = "未處理"
        elif "已處理" in txt:
            status = "已處理"
        elif "已完成" in txt:
            status = "已完成"

        purchase_status_name = ""
        if "待付款" in txt:
            purchase_status_name = "待付款"
        elif "已付款" in txt:
            purchase_status_name = "已付款"
        elif "取消訂單" in txt:
            purchase_status_name = "取消訂單"
        elif "已退款" in txt:
            purchase_status_name = "已退款"

        address = extract_address_from_text_block(txt)

        # 嘗試抓電話
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
            "address": address,
            "phone": phone,
            "edit_url": edit_url,
            "purchase_status_name": purchase_status_name,
            "purchase_status": "1" if purchase_status_name == "已付款" else "",
        })

    dedup = {}
    for item in data:
        dedup[item["order_no"]] = item
    return list(dedup.values())


# ========================
# 解析編輯頁（完整表單）
# ========================
def parse_edit_page(session, edit_url, phone=""):
    params = {}
    if phone:
        params["phone"] = phone

    r = session.get(edit_url, params=params, timeout=30)
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
            selected = el.select_one("option[selected]")
            if selected is not None:
                fields[name] = selected.get("value", "")
            else:
                first = el.select_one("option")
                fields[name] = first.get("value", "") if first else ""

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

    notice = ""
    notice_el = soup.select_one('textarea[name="notice"]')
    if notice_el:
        notice = notice_el.text.strip()
    else:
        notice = str(fields.get("notice", "")).strip()

    page_text = soup.get_text("\n", strip=True)

    order_no = ""
    m = re.search(r"(LC\d+)", page_text)
    if m:
        order_no = m.group(1)

    # 補抓姓名、地址、電話
    customer_name = str(fields.get("name", "") or fields.get("customer_name", "") or "").strip()
    phone_value = str(fields.get("phone", "")).strip()
    address_value = str(fields.get("address", "")).strip()
    progress = str(fields.get("progress", "")).strip()

    return {
        "action": action,
        "fields": fields,
        "notice": notice,
        "progress": progress,
        "order_no": order_no,
        "edit_url": r.url,
        "query_params": current_query_params,
        "customer_name": customer_name,
        "phone": phone_value,
        "address": address_value,
    }


# ========================
# 找前次 + 目標
# ========================
def find_previous_processed(current_order_no, current_address, current_phone, current_service_date, items):
    matched = []
    for x in items:
        if x["order_no"] == current_order_no:
            continue
        if x.get("purchase_status") != "1":
            continue
        if x["status"] not in ("已處理", "已完成"):
            continue
        if not x["date_obj"]:
            continue
        if current_service_date and x["date_obj"] >= current_service_date:
            continue
        if not same_address(x["address"], current_address):
            continue
        if current_phone and x.get("phone") and normalize_phone(x.get("phone")) != normalize_phone(current_phone):
            continue
        matched.append(x)

    if not matched:
        return None

    matched.sort(key=lambda k: k["date_obj"], reverse=True)
    return matched[0]


def find_current_unprocessed_same_address(current_order_no, current_address, current_phone, items):
    targets = []
    for x in items:
        if x["order_no"] != current_order_no:
            continue
        if x["status"] != "未處理":
            continue
        if not same_address(x["address"], current_address):
            continue
        if current_phone and x.get("phone") and normalize_phone(x.get("phone")) != normalize_phone(current_phone):
            continue
        targets.append(x)
    return targets


# ========================
# 送出 / 驗證
# ========================
def submit_update(session, form_info, phone, new_notice):
    action = form_info["action"]
    fields = dict(form_info["fields"])
    query_params = dict(form_info.get("query_params", {}))

    fields["notice"] = new_notice
    fields["progress"] = "1"

    if phone and "phone" not in query_params:
        query_params["phone"] = phone

    resp = session.post(
        action,
        params=query_params,
        files={k: (None, str(v)) for k, v in fields.items()},
        headers={
            "Referer": form_info["edit_url"],
            "User-Agent": "Mozilla/5.0",
        },
        timeout=30,
        allow_redirects=True,
    )
    resp.raise_for_status()
    return resp


def verify_update(session, edit_url, phone, expected_notice):
    form = parse_edit_page(session, edit_url, phone)
    actual_notice = str(form.get("notice", "")).strip()
    actual_progress = str(form.get("progress", "")).strip()

    notice_ok = actual_notice == str(expected_notice or "").strip()
    progress_ok = actual_progress == "1"

    return notice_ok and progress_ok, form


# ========================
# Sheet 樣式
# ========================
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
                "properties": {"pixelSize": 20},
                "fields": "pixelSize"
            }
        })

    requests_body.append({
        "repeatCell": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": 1,
                "startColumnIndex": 22,  # W
                "endColumnIndex": 24,    # X
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

    ws.spreadsheet.batch_update({"requests": requests_body})


# ========================
# log sheet
# ========================
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
    status: str,
    error_msg: str,
    full_log: str,
):
    log_ws.append_row([
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
        status,
        error_msg,
        clip_text(full_log, 20000),
    ])


# ========================
# 單筆處理核心（供兩種模式共用）
# ========================
def process_single_case(
    session,
    order,
    name,
    phone,
    addr,
    date,
    log,
):
    current_service_date = parse_date(date)

    log(f"訂單: {order}")
    log(f"客戶: {name}")
    log(f"電話: {phone}")
    log(f"地址: {addr}")
    log(f"日期: {date}")

    items = search_paid_orders_by_phone(session, phone)

    log("\n[列表頁候選]")
    for i in items:
        log(f"{i['order_no']} {i['date_str']} {i['status']} {i['address']}")

    prev = find_previous_processed(order, addr, phone, current_service_date, items)
    if not prev:
        log("❌ 沒有上一筆")
        return {
            "ok": False,
            "prev_date": "",
            "prev_order": "",
            "prev_notice": "",
            "error": "沒有上一筆",
        }

    targets = find_current_unprocessed_same_address(order, addr, phone, items)
    if not targets:
        log("❌ 沒有未處理目標單")
        return {
            "ok": False,
            "prev_date": prev["date_str"],
            "prev_order": prev["order_no"],
            "prev_notice": "",
            "error": "沒有未處理目標單",
        }

    log(f"\n[上一筆] {prev['order_no']} {prev['date_str']} {prev['status']} {prev['address']}")

    prev_form = parse_edit_page(session, prev["edit_url"], phone)
    prev_notice = str(prev_form.get("notice", "")).strip()

    if not prev_notice:
        log("❌ 上一筆找不到客服備註")
        return {
            "ok": False,
            "prev_date": prev["date_str"],
            "prev_order": prev["order_no"],
            "prev_notice": "",
            "error": "上一筆找不到客服備註",
        }

    submit_ok = 0
    verify_ok = 0
    failed_targets = []

    for t in targets:
        log(f"👉 寫入 {t['order_no']} {t['date_str']} {t['status']} {t['address']}")
        target_form = parse_edit_page(session, t["edit_url"], phone)

        try:
            submit_update(session, target_form, phone, prev_notice)
            time.sleep(SLEEP_SECONDS)
            submit_ok += 1

            ok, verified_form = verify_update(session, t["edit_url"], phone, prev_notice)
            time.sleep(SLEEP_SECONDS)

            if ok:
                verify_ok += 1
                log(f"✅ 驗證成功 {t['order_no']}")
            else:
                failed_targets.append(t["order_no"])
                log(
                    f"❌ 驗證失敗 {t['order_no']} "
                    f"(progress={verified_form.get('progress','')}, "
                    f"notice_head={str(verified_form.get('notice',''))[:30]})"
                )
        except Exception as e:
            failed_targets.append(t["order_no"])
            log(f"❌ 寫入失敗 {t['order_no']}：{e}")

    if verify_ok == len(targets):
        log(f"✅ 成功：已回填 {verify_ok} 筆")
        return {
            "ok": True,
            "prev_date": prev["date_str"],
            "prev_order": prev["order_no"],
            "prev_notice": prev_notice,
            "error": "",
        }

    log(
        f"❌ 失敗：已送出 {submit_ok} 筆，驗證成功 {verify_ok} 筆，"
        f"失敗目標：{', '.join(failed_targets)}"
    )
    return {
        "ok": False,
        "prev_date": prev["date_str"],
        "prev_order": prev["order_no"],
        "prev_notice": prev_notice,
        "error": f"已送出 {submit_ok} 筆，驗證成功 {verify_ok} 筆，失敗目標：{', '.join(failed_targets)}",
    }


# ========================
# Sheet 模式
# ========================
def main(row_spec="2", force=False, ui_logger=None):
    ws = get_ws()
    log_ws = get_log_ws()
    rows = ws.get_all_values()

    row_nums = parse_row_spec(row_spec)
    log = make_logger(ui_logger)

    sessions = {}
    updates = []
    updated_row_numbers = []

    result = {"processed": 0, "success": 0, "failed": 0, "skipped": 0, "errors": []}

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

            region = match_region_by_address(addr)
            if not region:
                log("❌ 無法判斷區域")
                result["failed"] += 1
                result["errors"].append(f"第{r}列：無法判斷區域")
                continue

            if region not in sessions:
                log(f"[登入] {region}")
                sessions[region] = login(region)

            s = sessions[region]

            single_result = process_single_case(
                session=s,
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
                status="成功" if single_result["ok"] else "失敗",
                error_msg=single_result["error"],
                full_log="\n".join(CURRENT_ROW_LOGS),
            )

            result["processed"] += 1
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
                status="失敗",
                error_msg=str(e),
                full_log="\n".join(CURRENT_ROW_LOGS),
            )
            result["failed"] += 1
            result["errors"].append(f"第{r}列：{e}")

    if updates:
        ws.batch_update(updates, value_input_option="RAW")
        apply_sheet_presentation(ws, updated_row_numbers)

    return result


# ========================
# 電話模式（不回寫主 sheet，但寫 log sheet）
# ========================
def main_by_phone(phone, ui_logger=None):
    CURRENT_ROW_LOGS.clear()
    log = make_logger(ui_logger)
    log_ws = get_log_ws()

    phone = normalize_phone(phone)
    result = {"processed": 0, "success": 0, "failed": 0, "skipped": 0, "errors": []}

    if not phone:
        msg = "電話號碼不可空白"
        log(f"❌ {msg}")
        result["failed"] = 1
        result["errors"].append(msg)
        return result

    log("\n===== BY電話 =====")
    log(f"輸入電話: {phone}")

    sessions = {}
    region_order = ["台北", "台中"]

    found_region = None
    found_session = None

    for region in region_order:
        try:
            log(f"[登入] {region}")
            sessions[region] = login(region)
            session = sessions[region]
            items = search_paid_orders_by_phone(session, phone)
            if items:
                found_region = region
                found_session = session
                break
        except Exception as e:
            log(f"❌ {region} 登入/搜尋失敗：{e}")

    if not found_session:
        msg = "查無已付款訂單"
        log(f"❌ {msg}")
        result["failed"] = 1
        result["errors"].append(msg)
        append_log_row(
            log_ws, "BY電話", phone, phone, "", "", "", "", "", "", "", "失敗", msg, "\n".join(CURRENT_ROW_LOGS)
        )
        return result

    items = search_paid_orders_by_phone(found_session, phone)
    log(f"[使用區域] {found_region}")
    log("\n[列表頁候選]")
    for i in items:
        log(f"{i['order_no']} {i['date_str']} {i['status']} {i['address']}")

    groups = find_phone_mode_targets(items)
    if not groups:
        msg = "找不到可處理的『未處理單 + 同地址前次已處理單』"
        log(f"❌ {msg}")
        result["failed"] = 1
        result["errors"].append(msg)
        append_log_row(
            log_ws, "BY電話", phone, phone, "", "", "", "", "", "", "", "失敗", msg, "\n".join(CURRENT_ROW_LOGS)
        )
        return result

    success_count = 0
    failed_count = 0

    for idx, group in enumerate(groups, start=1):
        target = group["target"]
        prev = group["prev"]

        log(f"\n===== 處理第 {idx} 組 =====")
        log(f"[目標單] {target['order_no']} {target['date_str']} {target['status']} {target['address']}")
        log(f"[上一筆] {prev['order_no']} {prev['date_str']} {prev['status']} {prev['address']}")

        try:
            prev_form = parse_edit_page(found_session, prev["edit_url"], phone)
            prev_notice = str(prev_form.get("notice", "")).strip()

            if not prev_notice:
                log("❌ 上一筆找不到客服備註")
                failed_count += 1
                append_log_row(
                    log_ws, "BY電話", phone, phone, "", target["address"], target["order_no"], target["date_str"],
                    prev["order_no"], prev["date_str"], "", "失敗", "上一筆找不到客服備註", "\n".join(CURRENT_ROW_LOGS)
                )
                continue

            target_form = parse_edit_page(found_session, target["edit_url"], phone)
            submit_update(found_session, target_form, phone, prev_notice)
            time.sleep(SLEEP_SECONDS)

            ok, verified_form = verify_update(found_session, target["edit_url"], phone, prev_notice)
            time.sleep(SLEEP_SECONDS)

            if ok:
                log(f"✅ 驗證成功 {target['order_no']}")
                success_count += 1
                append_log_row(
                    log_ws, "BY電話", phone, phone, verified_form.get("customer_name", ""), target["address"],
                    target["order_no"], target["date_str"], prev["order_no"], prev["date_str"], prev_notice,
                    "成功", "", "\n".join(CURRENT_ROW_LOGS)
                )
            else:
                log(
                    f"❌ 驗證失敗 {target['order_no']} "
                    f"(progress={verified_form.get('progress','')}, "
                    f"notice_head={str(verified_form.get('notice',''))[:30]})"
                )
                failed_count += 1
                append_log_row(
                    log_ws, "BY電話", phone, phone, verified_form.get("customer_name", ""), target["address"],
                    target["order_no"], target["date_str"], prev["order_no"], prev["date_str"], prev_notice,
                    "失敗", "送出後驗證失敗", "\n".join(CURRENT_ROW_LOGS)
                )

        except Exception as e:
            log(f"❌ 處理失敗 {target['order_no']}：{e}")
            failed_count += 1
            append_log_row(
                log_ws, "BY電話", phone, phone, "", target["address"], target["order_no"], target["date_str"],
                prev["order_no"], prev["date_str"], "", "失敗", str(e), "\n".join(CURRENT_ROW_LOGS)
            )

    result["processed"] = len(groups)
    result["success"] = success_count
    result["failed"] = failed_count

    if failed_count > 0:
        result["errors"].append(f"成功 {success_count} 筆，失敗 {failed_count} 筆")

    return result


# ========================
# 搜尋條件模式（不回寫主 sheet，但寫 log sheet）
# ========================
def main_by_conditions(date_s: str, limit: int = 10, ui_logger=None):
    CURRENT_ROW_LOGS.clear()
    log = make_logger(ui_logger)
    log_ws = get_log_ws()

    result = {"processed": 0, "success": 0, "failed": 0, "skipped": 0, "errors": []}

    if not parse_date(date_s):
        msg = "日期格式錯誤，請輸入 YYYY/MM/DD"
        log(f"❌ {msg}")
        result["failed"] = 1
        result["errors"].append(msg)
        return result

    log("\n===== BY搜尋條件 =====")
    log(f"訂購日期: {date_s}")
    log(f"條件: 付款狀態非取消 + 服務狀態未處理 + 每次只跑 {limit} 筆")

    sessions = {}
    region_order = ["台北", "台中"]

    all_targets = []

    for region in region_order:
        try:
            log(f"[登入] {region}")
            sessions[region] = login(region)
            session = sessions[region]
            items = search_by_conditions(session, date_s, limit=0)
            for item in items:
                item["region"] = region
            all_targets.extend(items)
        except Exception as e:
            log(f"❌ {region} 搜尋失敗：{e}")

    if not all_targets:
        msg = "查無符合條件訂單"
        log(f"❌ {msg}")
        result["failed"] = 1
        result["errors"].append(msg)
        append_log_row(
            log_ws, "BY搜尋條件", date_s, "", "", "", "", "", "", "", "", "失敗", msg, "\n".join(CURRENT_ROW_LOGS)
        )
        return result

    # 控制每次只跑 X 筆
    all_targets = all_targets[:limit]

    log("\n[搜尋結果]")
    for i in all_targets:
        log(f"{i['order_no']} {i['date_str']} {i['status']} {i['address']} {i['region']}")

    success_count = 0
    failed_count = 0

    for idx, target in enumerate(all_targets, start=1):
        try:
            CURRENT_ROW_LOGS.append("")  # 保留段落間隔
            log(f"\n===== 處理第 {idx} 筆 =====")
            log(f"[目標單] {target['order_no']} {target['date_str']} {target['status']} {target['address']} {target['region']}")

            session = sessions[target["region"]]
            target_form = parse_edit_page(session, target["edit_url"], target.get("phone", ""))

            target_phone = normalize_phone(target_form.get("phone", "") or target.get("phone", ""))
            target_name = target_form.get("customer_name", "")
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
                log(f"{i['order_no']} {i['date_str']} {i['status']} {i['address']}")

            prev = find_previous_processed(
                target["order_no"],
                target_addr,
                target_phone,
                target["date_obj"],
                items,
            )
            if not prev:
                raise RuntimeError("沒有上一筆已付款且相同電話地址的訂單")

            log(f"\n[上一筆] {prev['order_no']} {prev['date_str']} {prev['status']} {prev['address']}")

            prev_form = parse_edit_page(session, prev["edit_url"], target_phone)
            prev_notice = str(prev_form.get("notice", "")).strip()
            if not prev_notice:
                raise RuntimeError("上一筆找不到客服備註")

            submit_update(session, target_form, target_phone, prev_notice)
            time.sleep(SLEEP_SECONDS)

            ok, verified_form = verify_update(session, target["edit_url"], target_phone, prev_notice)
            time.sleep(SLEEP_SECONDS)

            if not ok:
                raise RuntimeError("送出後驗證失敗")

            log(f"✅ 驗證成功 {target['order_no']}")
            success_count += 1

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
                status="成功",
                error_msg="",
                full_log="\n".join(CURRENT_ROW_LOGS),
            )

        except Exception as e:
            failed_count += 1
            log(f"❌ 處理失敗 {target.get('order_no','')}：{e}")

            append_log_row(
                log_ws=log_ws,
                source_type="BY搜尋條件",
                source_value=date_s,
                phone=target.get("phone", ""),
                name="",
                address=target.get("address", ""),
                current_order=target.get("order_no", ""),
                current_date=target.get("date_str", ""),
                prev_order="",
                prev_date="",
                prev_notice="",
                status="失敗",
                error_msg=str(e),
                full_log="\n".join(CURRENT_ROW_LOGS),
            )

    result["processed"] = len(all_targets)
    result["success"] = success_count
    result["failed"] = failed_count

    if failed_count > 0:
        result["errors"].append(f"成功 {success_count} 筆，失敗 {failed_count} 筆")

    return result
