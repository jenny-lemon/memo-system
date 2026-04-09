# -*- coding: utf-8 -*-

import re
import time
from datetime import datetime
from urllib.parse import urljoin
from typing import Optional, List, Dict, Set, Callable

import gspread
import requests
from bs4 import BeautifulSoup
from google.oauth2.service_account import Credentials
from gspread.exceptions import APIError, SpreadsheetNotFound, WorksheetNotFound

import streamlit as st

import accounts
import env


# =========================
# 基本設定
# =========================
SHEET_ID = env.SHEET_ID
WORKSHEET_NAME = getattr(env, "WORKSHEET_NAME", "memo")
SLEEP_SECONDS = getattr(env, "SLEEP_SECONDS", 0.5)

ENV_NAME = getattr(env, "ENV", "prod").lower()
BASE_URL_DEV = getattr(env, "BASE_URL_DEV", "https://backend-dev.lemonclean.com.tw")
BASE_URL_PROD = getattr(env, "BASE_URL_PROD", "https://backend.lemonclean.com.tw")

if ENV_NAME == "dev":
    BASE_URL = BASE_URL_DEV.rstrip("/")
else:
    BASE_URL = BASE_URL_PROD.rstrip("/")

LOGIN_URL = f"{BASE_URL}/login"
PURCHASE_URL = f"{BASE_URL}/purchase"

LOG_BUFFER: List[str] = []


# =========================
# 工具
# =========================
def make_logger(ui_logger: Optional[Callable[[str], None]] = None):
    def _log(msg: str) -> None:
        print(msg, flush=True)
        LOG_BUFFER.append(msg)
        if ui_logger:
            ui_logger(msg)
    return _log


def normalize_text(text: str) -> str:
    if text is None:
        return ""
    text = str(text).replace("\u3000", " ")
    text = re.sub(r"\s+", "", text)
    return text.strip()


def normalize_phone(phone: str) -> str:
    return re.sub(r"\D+", "", str(phone or ""))


def parse_date(text: str):
    if not text:
        return None

    s = str(text).strip()
    patterns = [
        "%Y/%m/%d",
        "%Y-%m-%d",
        "%Y/%m/%d %H:%M:%S",
        "%Y-%m-%d %H:%M:%S",
        "%Y/%m/%d %H:%M",
        "%Y-%m-%d %H:%M",
    ]
    for fmt in patterns:
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass

    m = re.match(r"^\s*(\d{4})[/-](\d{1,2})[/-](\d{1,2})\s*$", s)
    if m:
        y, mo, d = map(int, m.groups())
        return datetime(y, mo, d)

    return None


def safe_cell(row_values, idx_1_based: int) -> str:
    i = idx_1_based - 1
    return str(row_values[i]).strip() if i < len(row_values) else ""


def build_absolute_url(href: str) -> str:
    href = (href or "").strip()
    if not href:
        return ""
    return urljoin(BASE_URL + "/", href)


def same_address(a: str, b: str) -> bool:
    na = normalize_text(a)
    nb = normalize_text(b)
    if not na or not nb:
        return False
    return na == nb or na in nb or nb in na


def extract_dates_from_text(text: str) -> List[datetime]:
    found = []
    for m in re.finditer(r"(\d{4}[/-]\d{1,2}[/-]\d{1,2})", str(text or "")):
        d = parse_date(m.group(1).replace("-", "/"))
        if d:
            found.append(d)
    return found


def looks_like_address_line(line: str) -> bool:
    if not line:
        return False
    keywords = ["市", "縣", "區", "鄉", "鎮", "路", "街", "段", "巷", "弄", "號"]
    return any(k in line for k in keywords)


def looks_like_address_continuation(line: str) -> bool:
    if not line:
        return False
    return any(k in line for k in ["樓", "室", "之", "A", "B", "C", "D"])


def extract_address_from_lines(lines: List[str]) -> str:
    clean = [x.strip() for x in lines if x and x.strip()]
    for i, line in enumerate(clean):
        if looks_like_address_line(line):
            addr = line
            if i + 1 < len(clean) and looks_like_address_continuation(clean[i + 1]):
                addr += clean[i + 1]
            return addr
    return ""


def parse_row_spec(spec: str) -> List[int]:
    rows: Set[int] = set()

    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue

        if "-" in part:
            a, b = part.split("-", 1)
            start = int(a.strip())
            end = int(b.strip())
            if start > end:
                start, end = end, start
            for n in range(start, end + 1):
                if n >= 2:
                    rows.add(n)
        else:
            n = int(part)
            if n >= 2:
                rows.add(n)

    return sorted(rows)


# =========================
# 區域判斷
# =========================
def match_region_by_address(address: str) -> Optional[str]:
    addr = str(address or "")
    for region, cfg in accounts.ACCOUNTS.items():
        for kw in cfg.get("address_keywords", []):
            if kw and kw in addr:
                return region
    return None


# =========================
# Google Sheet
# =========================
def get_google_worksheet():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    try:
        creds_dict = dict(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
    except Exception as e:
        raise RuntimeError(
            "讀不到 Streamlit Secrets 的 GOOGLE_SERVICE_ACCOUNT。"
            "請到 Streamlit Cloud → Manage app → Secrets 設定 [GOOGLE_SERVICE_ACCOUNT]。"
        ) from e

    try:
        creds = Credentials.from_service_account_info(
            creds_dict,
            scopes=scopes,
        )
        gc = gspread.authorize(creds)
        sh = gc.open_by_key(SHEET_ID)
        return sh.worksheet(WORKSHEET_NAME)
    except WorksheetNotFound as e:
        raise RuntimeError(f"找不到工作表 '{WORKSHEET_NAME}'。") from e
    except SpreadsheetNotFound as e:
        raise RuntimeError("找不到 Spreadsheet，請確認 SHEET_ID 是否正確。") from e
    except APIError as e:
        client_email = creds_dict.get("client_email", "(unknown)")
        raise RuntimeError(
            f"Google Sheet API 存取失敗。請確認這個 service account 已被分享進該 Sheet 並有編輯權限：{client_email}"
        ) from e


# =========================
# 後台
# =========================
def get_soup(session: requests.Session, url: str, params=None) -> BeautifulSoup:
    resp = session.get(url, params=params, timeout=30)
    resp.raise_for_status()
    return BeautifulSoup(resp.text, "html.parser")


def extract_csrf_token(soup: BeautifulSoup) -> str:
    el = soup.select_one('meta[name="csrf-token"]')
    if el and el.get("content"):
        return el["content"]

    el = soup.select_one('input[name="_token"]')
    if el and el.get("value"):
        return el["value"]

    return ""


def login_backend(region: str) -> requests.Session:
    cfg = accounts.ACCOUNTS[region]
    email = cfg.get("email", "").strip()
    password = cfg.get("password", "").strip()

    if not email or not password:
        raise RuntimeError(f"{region} 缺少 email/password")

    session = requests.Session()
    session.headers.update({
        "User-Agent": "Mozilla/5.0",
        "Accept-Language": "zh-TW,zh;q=0.9,en-US;q=0.8,en;q=0.7",
    })

    login_page = session.get(LOGIN_URL, timeout=30)
    login_page.raise_for_status()
    soup = BeautifulSoup(login_page.text, "html.parser")
    token = extract_csrf_token(soup)

    payload = {
        "_token": token,
        "email": email,
        "password": password,
    }

    resp = session.post(
        LOGIN_URL,
        data=payload,
        headers={"Referer": LOGIN_URL},
        timeout=30,
        allow_redirects=True,
    )
    resp.raise_for_status()

    check = session.get(PURCHASE_URL, timeout=30, allow_redirects=True)
    check.raise_for_status()
    if "/login" in check.url:
        raise RuntimeError(f"{region} 登入失敗")

    return session


def search_orders_by_phone(session: requests.Session, phone: str) -> List[Dict]:
    soup = get_soup(
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

    items = []
    rows = soup.select("table tbody tr")
    if not rows:
        rows = soup.select("tr")

    for tr in rows:
        edit_a = tr.select_one('a[href*="/purchase/edit/"]')
        if not edit_a:
            continue

        href = edit_a.get("href", "")
        m = re.search(r"/purchase/edit/(\d+)", href)
        if not m:
            continue

        tds = tr.find_all("td")
        row_text = tr.get_text("\n", strip=True)

        order_no = ""
        address = ""
        phone_in_row = ""
        service_date = None
        status = ""

        if len(tds) >= 5:
            td0_text = tds[0].get_text("\n", strip=True)
            td1_lines = [x.strip() for x in tds[1].get_text("\n", strip=True).splitlines() if x.strip()] if len(tds) > 1 else []
            td2_text = tds[2].get_text("\n", strip=True) if len(tds) > 2 else ""
            td4_text = tds[4].get_text("\n", strip=True) if len(tds) > 4 else ""

            m_order = re.search(r"(LC\d+)", td0_text)
            if m_order:
                order_no = m_order.group(1)

            for line in td1_lines:
                m_phone = re.search(r"(09\d{8})", line)
                if m_phone:
                    phone_in_row = m_phone.group(1)
                    break

            address = extract_address_from_lines(td1_lines)

            dates = extract_dates_from_text(td2_text)
            if dates:
                service_date = dates[0]

            if "未處理" in td4_text:
                status = "未處理"
            elif "已處理" in td4_text:
                status = "已處理"
            elif "已完成" in td4_text:
                status = "已完成"

        if not order_no:
            m_order = re.search(r"(LC\d+)", row_text)
            if m_order:
                order_no = m_order.group(1)

        if not service_date:
            dates = extract_dates_from_text(row_text)
            if dates:
                service_date = max(dates)

        if not status:
            if "未處理" in row_text:
                status = "未處理"
            elif "已處理" in row_text:
                status = "已處理"
            elif "已完成" in row_text:
                status = "已完成"

        if not address:
            lines = [x.strip() for x in row_text.splitlines() if x.strip()]
            address = extract_address_from_lines(lines)

        items.append({
            "purchase_id": m.group(1),
            "edit_url": build_absolute_url(href),
            "order_no": order_no,
            "service_date": service_date,
            "status": status,
            "address": address,
            "phone": phone_in_row or phone,
            "row_text": row_text,
        })

    dedup = {}
    for x in items:
        dedup[x["purchase_id"]] = x
    return list(dedup.values())


def parse_edit_form(session: requests.Session, edit_url: str, phone: str = "") -> Dict:
    params = {"phone": phone} if phone else None
    soup = get_soup(session, edit_url, params=params)

    form = soup.select_one("form")
    if not form:
        raise RuntimeError(f"找不到表單: {edit_url}")

    action = form.get("action") or edit_url
    action = build_absolute_url(action)

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

    page_text = soup.get_text("\n", strip=True)

    order_no = ""
    address = ""
    form_phone = ""
    service_date = None

    for k, v in fields.items():
        lk = k.lower()
        sv = str(v).strip()

        if not order_no and ("orderno" in lk or "order_no" in lk):
            order_no = sv

        if not address and ("address" in lk or "addr" in lk):
            if sv:
                address = sv

        if not form_phone and any(x in lk for x in ["phone", "mobile", "tel"]):
            if sv:
                form_phone = sv

        if not service_date and any(x in lk for x in ["clean_date", "service_date", "cleaning_date"]):
            parsed = parse_date(sv)
            if parsed:
                service_date = parsed

    if not order_no:
        m = re.search(r"(LC\d+)", page_text)
        if m:
            order_no = m.group(1)

    if not form_phone:
        m = re.search(r"(09\d{8})", page_text)
        if m:
            form_phone = m.group(1)

    if not service_date:
        for v in fields.values():
            parsed = parse_date(str(v).strip())
            if parsed:
                service_date = parsed
                break

    notice_el = soup.select_one('textarea[name="notice"]')
    notice_value = notice_el.text.strip() if notice_el else str(fields.get("notice", "")).strip()

    if "_token" not in fields:
        token = extract_csrf_token(soup)
        if token:
            fields["_token"] = token

    return {
        "action": action,
        "fields": fields,
        "order_no": order_no.strip(),
        "address": address.strip(),
        "phone": form_phone.strip(),
        "service_date": service_date,
        "notice": notice_value,
        "edit_url": edit_url,
    }


def find_previous_processed_order(current_order_no: str, current_address: str, current_phone: str, current_service_date, candidates: List[Dict]) -> Optional[Dict]:
    addr_now = normalize_text(current_address)
    phone_now = normalize_phone(current_phone)

    matched = []
    for item in candidates:
        if item.get("order_no") == current_order_no:
            continue
        if item.get("status") not in {"已處理", "已完成"}:
            continue
        if not item.get("service_date"):
            continue
        if current_service_date and item["service_date"] >= current_service_date:
            continue
        if normalize_phone(item.get("phone", "")) and normalize_phone(item.get("phone", "")) != phone_now:
            continue
        if not same_address(item.get("address", ""), addr_now):
            continue
        matched.append(item)

    if not matched:
        return None

    matched.sort(key=lambda x: x["service_date"], reverse=True)
    return matched[0]


def find_unprocessed_targets(current_address: str, current_phone: str, candidates: List[Dict]) -> List[Dict]:
    addr_now = normalize_text(current_address)
    phone_now = normalize_phone(current_phone)

    targets = []
    for item in candidates:
        if item.get("status") != "未處理":
            continue
        if normalize_phone(item.get("phone", "")) and normalize_phone(item.get("phone", "")) != phone_now:
            continue
        if not same_address(item.get("address", ""), addr_now):
            continue
        targets.append(item)

    targets.sort(key=lambda x: (x.get("service_date") or datetime.max, x.get("order_no", "")))
    return targets


def submit_update_processed_with_notice(session: requests.Session, form_info: Dict, phone: str, new_notice: str):
    action = form_info["action"]
    fields = dict(form_info["fields"])
    fields["notice"] = new_notice
    fields["progress"] = "1"

    params = {"phone": phone} if phone else None

    resp = session.post(
        action,
        params=params,
        files={k: (None, str(v)) for k, v in fields.items()},
        headers={"Referer": f"{form_info['edit_url']}?phone={phone}" if phone else form_info["edit_url"]},
        timeout=30,
        allow_redirects=True,
    )
    resp.raise_for_status()
    return resp


def main(row_spec: Optional[str] = None, force: bool = False, ui_logger: Optional[Callable[[str], None]] = None):
    logger = make_logger(ui_logger)

    logger(f"[環境] ENV={ENV_NAME} | BASE_URL={BASE_URL}")
    worksheet = get_google_worksheet()

    all_values = worksheet.get_all_values()
    max_row = len(all_values)

    stats = {
        "processed": 0,
        "success": 0,
        "failed": 0,
        "skipped": 0,
        "errors": [],
    }

    if max_row < 2:
        logger("sheet 沒資料")
        return stats

    if row_spec:
        row_numbers = [r for r in parse_row_spec(row_spec) if r <= max_row]
        logger(f"[模式] 指定列：{row_spec} -> {row_numbers}")
    else:
        row_numbers = list(range(2, max_row + 1))
        logger("[模式] 全部列")

    if force:
        logger("[選項] 強制重跑：忽略 V 欄已有值")

    sessions_by_region = {}
    phone_search_cache: Dict[str, List[Dict]] = {}
    detail_cache: Dict[str, Dict] = {}
    pending_updates = []

    for row_num in row_numbers:
        LOG_BUFFER.clear()

        try:
            row_values = all_values[row_num - 1]

            logger(f"\n===== 第 {row_num} 列 =====")
            logger(f"原始列資料: {row_values}")

            current_order_no = safe_cell(row_values, 2)
            service_date_str = safe_cell(row_values, 8)
            current_address = safe_cell(row_values, 14)
            current_phone = safe_cell(row_values, 15)
            sheet_status = safe_cell(row_values, 22)

            logger(f"訂單編號: {current_order_no}")
            logger(f"服務日期: {service_date_str}")
            logger(f"地址: {current_address}")
            logger(f"電話: {current_phone}")
            logger(f"V欄: {sheet_status}")

            if sheet_status and not force:
                logger(f"[SKIP] 第{row_num}列 已處理 ({sheet_status})")
                stats["skipped"] += 1
                continue

            if not current_order_no or not current_address or not current_phone:
                logger(f"[SKIP] 第{row_num}列 缺少訂單編號/地址/電話")
                stats["skipped"] += 1
                continue

            current_service_date = parse_date(service_date_str)
            if not current_service_date:
                logger(f"[SKIP] 第{row_num}列 服務日期格式錯誤: {service_date_str}")
                stats["skipped"] += 1
                continue

            region = match_region_by_address(current_address)
            if not region:
                logger(f"[SKIP] 第{row_num}列 無法依地址判斷區域: {current_address}")
                stats["skipped"] += 1
                continue

            if region not in sessions_by_region:
                logger(f"[登入] {region}")
                sessions_by_region[region] = login_backend(region)
                time.sleep(SLEEP_SECONDS)

            session = sessions_by_region[region]
            phone_key = f"{region}|{normalize_phone(current_phone)}"

            logger(f"[處理] 第{row_num}列 | 區域={region} | 訂單={current_order_no} | 電話={current_phone}")
            stats["processed"] += 1

            if phone_key not in phone_search_cache:
                phone_search_cache[phone_key] = search_orders_by_phone(session, normalize_phone(current_phone))
                time.sleep(SLEEP_SECONDS)

            search_results = phone_search_cache[phone_key]
            if not search_results:
                logger(f"[略過] 第{row_num}列 查無同電話訂單，不改動")
                stats["skipped"] += 1
                continue

            logger("[列表頁候選]")
            for item in search_results:
                date_str = item.get("service_date").strftime("%Y/%m/%d") if item.get("service_date") else ""
                logger(f"{item.get('order_no', '')} {date_str} {item.get('status', '')} {item.get('address', '')}")

            prev_brief = find_previous_processed_order(
                current_order_no=current_order_no,
                current_address=current_address,
                current_phone=current_phone,
                current_service_date=current_service_date,
                candidates=search_results,
            )

            if not prev_brief:
                logger(f"[略過] 第{row_num}列 沒有上次訂單，不改動 S/T/U/V")
                stats["skipped"] += 1
                continue

            targets_brief = find_unprocessed_targets(
                current_address=current_address,
                current_phone=current_phone,
                candidates=search_results,
            )

            if not targets_brief:
                logger(f"[略過] 第{row_num}列 沒有同地址未處理訂單可回填")
                stats["skipped"] += 1
                continue

            prev_form = None
            target_forms = []
            need_order_nos = {prev_brief.get("order_no")} | {x.get("order_no") for x in targets_brief}

            for item in search_results:
                if item.get("order_no") not in need_order_nos:
                    continue

                cache_key = f"{region}|{item['purchase_id']}"
                if cache_key not in detail_cache:
                    detail_cache[cache_key] = parse_edit_form(
                        session,
                        item["edit_url"],
                        normalize_phone(current_phone),
                    )
                    time.sleep(SLEEP_SECONDS)

                detail = detail_cache[cache_key]

                if detail.get("order_no") == prev_brief.get("order_no"):
                    prev_form = detail

                if detail.get("order_no") in {x.get("order_no") for x in targets_brief}:
                    target_forms.append(detail)

            if not prev_form:
                logger(f"[失敗] 第{row_num}列 找不到上一筆訂單編輯頁")
                stats["failed"] += 1
                stats["errors"].append(f"第{row_num}列：找不到上一筆訂單編輯頁")
                pending_updates.append({
                    "range": f"V{row_num}:W{row_num}",
                    "values": [["失敗", "\n".join(LOG_BUFFER)]],
                })
                continue

            prev_notice = prev_form.get("notice", "").strip()
            prev_service_date = prev_brief["service_date"].strftime("%Y/%m/%d")
            prev_order_no = prev_brief.get("order_no", "").strip()

            updated_count = 0
            updated_order_nos = []

            for target_form in target_forms:
                submit_update_processed_with_notice(
                    session=session,
                    form_info=target_form,
                    phone=normalize_phone(current_phone),
                    new_notice=prev_notice,
                )
                updated_count += 1
                updated_order_nos.append(target_form.get("order_no", ""))
                time.sleep(SLEEP_SECONDS)

            logger(
                f"[成功] 第{row_num}列 -> 已回填並改成已處理 {updated_count} 筆；"
                f"上次日期 {prev_service_date} / 上次單號 {prev_order_no} / 目標 {', '.join(updated_order_nos)}"
            )

            pending_updates.append({
                "range": f"S{row_num}:W{row_num}",
                "values": [[
                    prev_service_date,
                    prev_order_no,
                    prev_notice,
                    "成功",
                    "\n".join(LOG_BUFFER)
                ]],
            })
            stats["success"] += 1

        except Exception as e:
            msg = f"[失敗] 第{row_num}列: {e}"
            logger(msg)
            stats["failed"] += 1
            stats["errors"].append(msg)
            pending_updates.append({
                "range": f"V{row_num}:W{row_num}",
                "values": [["失敗", "\n".join(LOG_BUFFER)]],
            })

    if pending_updates:
        worksheet.batch_update(pending_updates, value_input_option="RAW")
        logger(f"[完成] 已批次寫回 {len(pending_updates)} 筆 sheet 更新")

    return stats
