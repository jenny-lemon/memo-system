# -*- coding: utf-8 -*-
import re
import time
from datetime import datetime
from urllib.parse import urljoin
from typing import Optional, List, Dict, Set, Callable

import requests
from bs4 import BeautifulSoup
import gspread
from google.oauth2.service_account import Credentials
import streamlit as st

import accounts
import env


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
SLEEP_SECONDS = getattr(env, "SLEEP_SECONDS", 0.5)
CURRENT_ROW_LOGS: List[str] = []


def get_ws():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    try:
        creds = Credentials.from_service_account_info(
            dict(st.secrets["GOOGLE_SERVICE_ACCOUNT"]),
            scopes=scopes,
        )
    except Exception:
        creds = Credentials.from_service_account_file(
            "service_account.json",
            scopes=scopes,
        )

    gc = gspread.authorize(creds)
    return gc.open_by_key(env.SHEET_ID).worksheet(WORKSHEET_NAME)


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


def match_region_by_address(address: str) -> Optional[str]:
    addr = str(address or "")
    for region, cfg in accounts.ACCOUNTS.items():
        for kw in cfg.get("address_keywords", []):
            if kw and kw in addr:
                return region
    return None


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
# 列表搜尋：手機 + 已付款
# ========================
def search_paid_orders_by_phone(session, phone):
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
            "purchase_status": "1",  # 已付款
            "progress_status": "",
            "invoiceStatus": "",
            "otherFee": "",
            "orderBy": "",
        },
        timeout=30,
    )
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "html.parser")

    data = []

    for tr in soup.select("tr"):
        txt = tr.get_text("\n", strip=True)

        m = re.search(r"(LC\d+)", txt)
        if not m:
            continue

        order_no = m.group(1)

        date_obj = None
        date_match = re.search(r"(\d{4}-\d{2}-\d{2})", txt)
        if date_match:
            date_obj = parse_date(date_match.group(1))

        status = ""
        if "未處理" in txt:
            status = "未處理"
        elif "已處理" in txt:
            status = "已處理"
        elif "已完成" in txt:
            status = "已完成"

        address = ""
        for line in txt.splitlines():
            s = line.strip()
            if any(k in s for k in ["市", "區", "路", "街", "段", "巷", "弄", "號"]):
                if len(s) >= 6 and "付款" not in s and "服務" not in s:
                    address = s
                    break

        edit_link = tr.select_one('a[href*="/purchase/edit/"]')
        edit_url = build_absolute_url(edit_link["href"]) if edit_link else ""

        data.append({
            "order_no": order_no,
            "date_str": date_match.group(1).replace("-", "/") if date_match else "",
            "date_obj": date_obj,
            "status": status,
            "address": address,
            "edit_url": edit_url,
        })

    dedup = {}
    for item in data:
        dedup[item["order_no"]] = item
    return list(dedup.values())


# ========================
# 解析編輯頁
# ========================
def parse_edit_page(session, edit_url, phone=""):
    params = {"phone": phone} if phone else None
    r = session.get(edit_url, params=params, timeout=30)
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "html.parser")

    form = soup.select_one("form")
    if not form:
        raise RuntimeError(f"找不到表單: {edit_url}")

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

    notice = ""
    notice_el = soup.select_one("textarea[name=notice]")
    if notice_el:
        notice = notice_el.text.strip()
    elif "notice" in fields:
        notice = str(fields.get("notice", "")).strip()

    order_no = ""
    text_all = soup.get_text("\n", strip=True)
    m = re.search(r"(LC\d+)", text_all)
    if m:
        order_no = m.group(1)

    progress = str(fields.get("progress", "")).strip()
    action = build_absolute_url(form.get("action") or edit_url)

    return {
        "action": action,
        "fields": fields,
        "notice": notice,
        "progress": progress,
        "order_no": order_no,
        "edit_url": edit_url,
    }


# ========================
# 找前次訂單
# ========================
def find_previous_processed(current_order_no, current_address, current_phone, current_service_date, items):
    addr_now = normalize_text(current_address)
    phone_now = normalize_phone(current_phone)

    matched = []
    for x in items:
        if x["order_no"] == current_order_no:
            continue
        if x["status"] not in ("已處理", "已完成"):
            continue
        if not x["date_obj"]:
            continue
        if current_service_date and x["date_obj"] >= current_service_date:
            continue
        if not same_address(x["address"], current_address):
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
        targets.append(x)
    return targets


# ========================
# 送出回寫
# ========================
def submit_update(session, form_info, phone, new_notice):
    action = form_info["action"]
    fields = dict(form_info["fields"])
    fields["notice"] = new_notice
    fields["progress"] = "1"

    resp = session.post(
        action,
        params={"phone": phone} if phone else None,
        files={k: (None, str(v)) for k, v in fields.items()},
        headers={"Referer": form_info["edit_url"]},
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
# Sheet 呈現：W/X 不撐高
# ========================
def apply_sheet_presentation(ws, updated_rows: List[int]):
    if not updated_rows:
        return

    sheet_id = ws._properties["sheetId"]

    requests_body = []

    # 設列高 20
    for row_num in updated_rows:
        requests_body.append({
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "ROWS",
                    "startIndex": row_num - 1,
                    "endIndex": row_num,
                },
                "properties": {
                    "pixelSize": 20
                },
                "fields": "pixelSize"
            }
        })

    # W:X 文字裁切
    # W=23, X=24 -> 0-based start 22, end 24
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

    ws.spreadsheet.batch_update({"requests": requests_body})


# ========================
# 主流程
# ========================
def main(row_spec="2", force=False, ui_logger=None):
    ws = get_ws()
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
                updates.append({
                    "range": f"V{r}:X{r}",
                    "values": [["失敗", "\n".join(CURRENT_ROW_LOGS), ""]],
                })
                updated_row_numbers.append(r)
                continue

            row = rows[r - 1]

            order = safe_cell(row, 2)       # B
            date = safe_cell(row, 8)        # H
            name = safe_cell(row, 13)       # M
            addr = safe_cell(row, 14)       # N
            phone = safe_cell(row, 15)      # O
            v_status = safe_cell(row, 22)   # V

            current_service_date = parse_date(date)

            log(f"\n===== 第{r}列 =====")
            log(f"訂單: {order}")
            log(f"客戶: {name}")
            log(f"電話: {phone}")
            log(f"地址: {addr}")
            log(f"日期: {date}")
            log(f"V欄: {v_status}")

            if v_status and not force:
                log("⏭ 已有狀態，略過")
                result["skipped"] += 1
                continue

            if not order or not addr or not phone:
                log("❌ 缺少訂單 / 地址 / 電話")
                result["failed"] += 1
                result["errors"].append(f"第{r}列：缺少訂單 / 地址 / 電話")
                updates.append({
                    "range": f"V{r}:X{r}",
                    "values": [["失敗", "\n".join(CURRENT_ROW_LOGS), ""]],
                })
                updated_row_numbers.append(r)
                continue

            region = match_region_by_address(addr)
            if not region:
                log("❌ 無法判斷區域")
                result["failed"] += 1
                result["errors"].append(f"第{r}列：無法判斷區域")
                updates.append({
                    "range": f"V{r}:X{r}",
                    "values": [["失敗", "\n".join(CURRENT_ROW_LOGS), ""]],
                })
                updated_row_numbers.append(r)
                continue

            if region not in sessions:
                log(f"[登入] {region}")
                sessions[region] = login(region)

            s = sessions[region]

            items = search_paid_orders_by_phone(s, phone)

            log("\n[列表頁候選]")
            for i in items:
                log(f"{i['order_no']} {i['date_str']} {i['status']} {i['address']}")

            prev = find_previous_processed(order, addr, phone, current_service_date, items)
            if not prev:
                log("❌ 沒有上一筆")
                result["failed"] += 1
                result["errors"].append(f"第{r}列：沒有上一筆")
                updates.append({
                    "range": f"V{r}:X{r}",
                    "values": [["失敗", "\n".join(CURRENT_ROW_LOGS), ""]],
                })
                updated_row_numbers.append(r)
                continue

            targets = find_current_unprocessed_same_address(order, addr, phone, items)
            if not targets:
                log("❌ 沒有未處理目標單")
                result["failed"] += 1
                result["errors"].append(f"第{r}列：沒有未處理目標單")
                updates.append({
                    "range": f"V{r}:X{r}",
                    "values": [["失敗", "\n".join(CURRENT_ROW_LOGS), ""]],
                })
                updated_row_numbers.append(r)
                continue

            log(f"\n[上一筆] {prev['order_no']} {prev['date_str']} {prev['status']} {prev['address']}")

            prev_form = parse_edit_page(s, prev["edit_url"], phone)
            prev_notice = str(prev_form.get("notice", "")).strip()

            if not prev_notice:
                log("❌ 上一筆找不到客服備註")
                result["failed"] += 1
                result["errors"].append(f"第{r}列：上一筆找不到客服備註")
                updates.append({
                    "range": f"S{r}:X{r}",
                    "values": [[
                        prev["date_str"],
                        prev["order_no"],
                        "",
                        "失敗",
                        "\n".join(CURRENT_ROW_LOGS),
                        ""
                    ]],
                })
                updated_row_numbers.append(r)
                continue

            submit_ok = 0
            verify_ok = 0
            failed_targets = []

            for t in targets:
                log(f"👉 寫入 {t['order_no']} {t['date_str']} {t['status']} {t['address']}")
                target_form = parse_edit_page(s, t["edit_url"], phone)

                try:
                    submit_update(s, target_form, phone, prev_notice)
                    time.sleep(SLEEP_SECONDS)
                    submit_ok += 1

                    ok, verified_form = verify_update(s, t["edit_url"], phone, prev_notice)
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

            # 只有全部 target 驗證成功，才算成功
            if verify_ok == len(targets):
                log(f"✅ 成功：已回填 {verify_ok} 筆")

                updates.append({
                    "range": f"S{r}:X{r}",
                    "values": [[
                        prev["date_str"],
                        prev["order_no"],
                        prev_notice,
                        "成功",
                        "\n".join(CURRENT_ROW_LOGS),
                        prev_notice
                    ]],
                })
                updated_row_numbers.append(r)
                result["success"] += 1
                result["processed"] += 1
            else:
                log(
                    f"❌ 失敗：已送出 {submit_ok} 筆，驗證成功 {verify_ok} 筆，"
                    f"失敗目標：{', '.join(failed_targets)}"
                )
                updates.append({
                    "range": f"S{r}:X{r}",
                    "values": [[
                        prev["date_str"],
                        prev["order_no"],
                        prev_notice,
                        "失敗",
                        "\n".join(CURRENT_ROW_LOGS),
                        prev_notice
                    ]],
                })
                updated_row_numbers.append(r)
                result["failed"] += 1
                result["processed"] += 1
                result["errors"].append(
                    f"第{r}列：已送出 {submit_ok} 筆，驗證成功 {verify_ok} 筆，失敗目標：{', '.join(failed_targets)}"
                )

        except Exception as e:
            log(f"❌ 例外錯誤：{e}")
            updates.append({
                "range": f"V{r}:X{r}",
                "values": [["失敗", "\n".join(CURRENT_ROW_LOGS), ""]],
            })
            updated_row_numbers.append(r)
            result["failed"] += 1
            result["errors"].append(f"第{r}列：{e}")

    if updates:
        ws.batch_update(updates, value_input_option="RAW")
        apply_sheet_presentation(ws, updated_row_numbers)

    return result
