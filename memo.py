# -*- coding: utf-8 -*-
import re
import time
from datetime import datetime
from urllib.parse import urljoin

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
ENV_NAME = "prod"

BASE_URL_DEV = env.BASE_URL_DEV
BASE_URL_PROD = env.BASE_URL_PROD

BASE_URL = BASE_URL_PROD
LOGIN_URL = ""
PURCHASE_URL = ""

def set_env(env_name):
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
def get_ws():
    creds = Credentials.from_service_account_info(
        dict(st.secrets["GOOGLE_SERVICE_ACCOUNT"]),
        scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    return gspread.authorize(creds).open_by_key(env.SHEET_ID).worksheet("memo")

# ========================
# 工具
# ========================
def normalize_phone(p):
    return re.sub(r"\D+", "", str(p))

def parse_date(t):
    try:
        return datetime.strptime(t, "%Y/%m/%d")
    except:
        return None

def parse_row_spec(spec):
    rows = set()
    for p in spec.split(","):
        if "-" in p:
            a, b = map(int, p.split("-"))
            rows.update(range(a, b+1))
        else:
            rows.add(int(p))
    return sorted(rows)

# ========================
# LOGIN
# ========================
def login(region):
    s = requests.Session()
    r = s.get(LOGIN_URL)
    soup = BeautifulSoup(r.text, "html.parser")
    token = soup.select_one("input[name=_token]").get("value")

    s.post(LOGIN_URL, data={
        "_token": token,
        "email": accounts.ACCOUNTS[region]["email"],
        "password": accounts.ACCOUNTS[region]["password"]
    })
    return s

# ========================
# 搜尋
# ========================
def search(session, phone):
    r = session.get(PURCHASE_URL, params={"phone": phone})
    soup = BeautifulSoup(r.text, "html.parser")

    data = []

    for tr in soup.select("tr"):
        txt = tr.get_text(" ", strip=True)

        m = re.search(r"(LC\d+)", txt)
        if not m:
            continue

        order_no = m.group(1)
        date = re.search(r"\d{4}-\d{2}-\d{2}", txt)
        status = "未處理" if "未處理" in txt else "已處理"

        address = ""
        for part in txt.split():
            if "市" in part:
                address = part

        data.append({
            "order_no": order_no,
            "date": date.group(0) if date else "",
            "status": status,
            "address": address,
        })

    return data

# ========================
# 主流程
# ========================
def main(row_spec="2", force=False, ui_logger=None):
    ws = get_ws()
    rows = ws.get_all_values()

    row_nums = parse_row_spec(row_spec)

    logs = []
    def log(msg):
        print(msg)
        logs.append(msg)
        if ui_logger:
            ui_logger(msg)

    sessions = {}
    updates = []

    result = {"processed":0,"success":0,"failed":0,"skipped":0}

    for r in row_nums:
        logs.clear()

        row = rows[r-1]

        order = row[1]
        date = row[7]
        name = row[12]
        addr = row[13]
        phone = row[14]

        log(f"\n===== 第{r}列 =====")
        log(f"訂單: {order}")
        log(f"客戶: {name}")
        log(f"電話: {phone}")
        log(f"地址: {addr}")
        log(f"日期: {date}")

        region = "台北"

        if region not in sessions:
            sessions[region] = login(region)

        s = sessions[region]

        items = search(s, phone)

        log("\n[列表頁候選]")
        for i in items:
            log(f"{i['order_no']} {i['date']} {i['status']} {i['address']}")

        prev = next((x for x in items if x["status"]=="已處理"), None)

        if not prev:
            log("❌ 沒有上一筆")
            updates.append((r,"失敗","\n".join(logs)))
            result["failed"]+=1
            continue

        targets = [x for x in items if x["status"]=="未處理"]

        if not targets:
            log("❌ 沒有未處理")
            updates.append((r,"失敗","\n".join(logs)))
            result["failed"]+=1
            continue

        log(f"\n[上一筆] {prev['order_no']} {prev['date']} {prev['address']}")

        for t in targets:
            log(f"👉 寫入 {t['order_no']}")

        updates.append((r,"成功","\n".join(logs)))
        result["success"]+=1
        result["processed"]+=1

    # 批次寫
    batch=[]
    for r,status,logtext in updates:
        batch.append({
            "range":f"V{r}:W{r}",
            "values":[[status,logtext]]
        })

    ws.batch_update(batch)

    return result
