# -*- coding: utf-8 -*-
import os
import re
import time
import json
from datetime import datetime
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup
import gspread
from google.oauth2.service_account import Credentials

import accounts

# ========================
# 環境設定
# ========================
ENV_NAME = "prod"

BASE_URL_DEV = "https://backend-dev.lemonclean.com.tw"
BASE_URL_PROD = "https://backend.lemonclean.com.tw"

BASE_URL = BASE_URL_PROD
LOGIN_URL = f"{BASE_URL}/login"
PURCHASE_URL = f"{BASE_URL}/purchase"

# ========================
# Google 設定
# ========================
SHEET_ID = "1de41gNvBZCGdfy0qNouRNEaQD7R019VAvz2cfq88ZrE"
WORKSHEET_NAME = "memo"


# ========================
# 工具
# ========================
def log(msg, ui_logger=None):
    print(msg)
    if ui_logger:
        ui_logger(msg)


def normalize_text(text):
    return re.sub(r"\s+", "", str(text or ""))


def normalize_phone(phone):
    return re.sub(r"\D+", "", str(phone or ""))


def parse_date(text):
    if not text:
        return None
    for fmt in ["%Y/%m/%d", "%Y-%m-%d"]:
        try:
            return datetime.strptime(text.strip(), fmt)
        except:
            pass
    return None


# ========================
# Google Sheet
# ========================
def get_google_worksheet():
    try:
        import streamlit as st

        creds_dict = st.secrets["GOOGLE_SERVICE_ACCOUNT"]
        creds = Credentials.from_service_account_info(
            creds_dict,
            scopes=[
                "https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive",
            ],
        )
    except:
        creds = Credentials.from_service_account_file(
            "service_account.json",
            scopes=[
                "https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive",
            ],
        )

    gc = gspread.authorize(creds)
    return gc.open_by_key(SHEET_ID).worksheet(WORKSHEET_NAME)


# ========================
# 解析列
# ========================
def parse_row_spec(spec):
    rows = set()
    parts = spec.split(",")
    for p in parts:
        if "-" in p:
            a, b = p.split("-")
            for i in range(int(a), int(b) + 1):
                rows.add(i)
        else:
            rows.add(int(p))
    return sorted(rows)


# ========================
# 登入
# ========================
def login(region):
    cfg = accounts.ACCOUNTS[region]

    email = cfg.get("email")
    password = cfg.get("password")

    if not email or not password:
        raise Exception(f"{region} 缺少 email/password")

    s = requests.Session()
    r = s.get(LOGIN_URL)
    soup = BeautifulSoup(r.text, "html.parser")

    token = soup.select_one("input[name=_token]").get("value")

    s.post(LOGIN_URL, data={
        "_token": token,
        "email": email,
        "password": password
    })

    return s


# ========================
# 搜尋列表頁
# ========================
def search_orders(session, phone):
    r = session.get(PURCHASE_URL, params={"phone": phone})
    soup = BeautifulSoup(r.text, "html.parser")

    results = []

    for row in soup.select("tr"):
        text = row.get_text(" ", strip=True)

        m = re.search(r"(LC\d+)", text)
        if not m:
            continue

        order_no = m.group(1)

        date_match = re.search(r"\d{4}-\d{2}-\d{2}", text)
        status = "未處理" if "未處理" in text else "已處理"

        results.append({
            "order_no": order_no,
            "date": date_match.group(0) if date_match else "",
            "status": status,
            "row_html": str(row)
        })

    return results


# ========================
# 取備註
# ========================
def get_notice(session, order_no):
    r = session.get(f"{PURCHASE_URL}", params={"keyword": order_no})
    soup = BeautifulSoup(r.text, "html.parser")

    edit_link = soup.select_one('a[href*="/purchase/edit/"]')
    if not edit_link:
        return ""

    edit_url = urljoin(BASE_URL, edit_link["href"])

    r2 = session.get(edit_url)
    soup2 = BeautifulSoup(r2.text, "html.parser")

    textarea = soup2.select_one("textarea[name=notice]")

    return textarea.text.strip() if textarea else ""


# ========================
# 寫回（改成已處理）
# ========================
def update_order(session, order_no, notice):
    r = session.get(f"{PURCHASE_URL}", params={"keyword": order_no})
    soup = BeautifulSoup(r.text, "html.parser")

    edit_link = soup.select_one('a[href*="/purchase/edit/"]')
    if not edit_link:
        return False

    edit_url = urljoin(BASE_URL, edit_link["href"])

    r2 = session.get(edit_url)
    soup2 = BeautifulSoup(r2.text, "html.parser")

    token = soup2.select_one("input[name=_token]").get("value")

    data = {}
    for inp in soup2.select("input, textarea"):
        name = inp.get("name")
        if not name:
            continue
        data[name] = inp.get("value", "")

    data["notice"] = notice
    data["progress"] = "1"  # 已處理

    session.post(edit_url, data=data)

    return True


# ========================
# 主流程
# ========================
def main(row_spec="2", force=False, ui_logger=None):
    ws = get_google_worksheet()
    all_rows = ws.get_all_values()

    target_rows = parse_row_spec(row_spec)

    sessions = {}

    updates = []
    result = {"processed": 0, "success": 0, "failed": 0, "skipped": 0, "errors": []}

    for row_num in target_rows:
        try:
            row = all_rows[row_num - 1]

            order_no = row[1]
            date = row[7]
            address = row[13]
            phone = row[14]
            status = row[21] if len(row) > 21 else ""

            log(f"\n===== 第{row_num}列 =====", ui_logger)

            if status and not force:
                log("[略過] 已有狀態", ui_logger)
                result["skipped"] += 1
                continue

            region = "台北" if "台北" in address or "新北" in address else "台中"

            if region not in sessions:
                log(f"[登入] {region}", ui_logger)
                sessions[region] = login(region)

            session = sessions[region]

            results = search_orders(session, phone)

            log("[列表頁候選]", ui_logger)
            for r in results:
                log(f"{r['order_no']} {r['date']} {r['status']}", ui_logger)

            prev = None
            for r in results:
                if r["status"] == "已處理" and r["order_no"] != order_no:
                    prev = r
                    break

            if not prev:
                log("[失敗] 找不到上一筆", ui_logger)
                updates.append([row_num, "", "", "", "失敗", "找不到上一筆"])
                result["failed"] += 1
                continue

            notice = get_notice(session, prev["order_no"])

            success_count = 0

            for r in results:
                if r["status"] == "未處理":
                    ok = update_order(session, r["order_no"], notice)
                    if ok:
                        success_count += 1

            if success_count > 0:
                updates.append([
                    row_num,
                    prev["date"].replace("-", "/"),
                    prev["order_no"],
                    notice,
                    "成功",
                    f"已回填 {success_count} 筆"
                ])
                result["success"] += 1
                log(f"[成功] 已回填 {success_count} 筆", ui_logger)
            else:
                updates.append([row_num, "", "", "", "失敗", "沒有成功寫入"])
                result["failed"] += 1

            result["processed"] += 1

        except Exception as e:
            log(f"[失敗] {e}", ui_logger)
            updates.append([row_num, "", "", "", "失敗", str(e)])
            result["failed"] += 1
            result["errors"].append(str(e))

    # 批次寫入（避免 429）
    if updates:
        batch = []
        for u in updates:
            r = u[0]
            batch.append({
                "range": f"S{r}:W{r}",
                "values": [u[1:]]
            })

        ws.batch_update(batch)

    log(f"\n[完成] 已批次寫回 {len(updates)} 筆", ui_logger)

    return result
