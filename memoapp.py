# -*- coding: utf-8 -*-
import streamlit as st
import memo
import accounts

st.set_page_config(page_title="Memo 自動回填系統", layout="wide")

st.title("📋 Memo 自動回填系統")

# ========================
# session state
# ========================
if "logs" not in st.session_state:
    st.session_state.logs = []

if "last_result" not in st.session_state:
    st.session_state.last_result = None


# ========================
# logger
# ========================
log_box = st.empty()


def render_logs():
    if st.session_state.logs:
        log_box.code("\n".join(st.session_state.logs[-1000:]))
    else:
        log_box.code("尚未執行")


def ui_log(msg):
    st.session_state.logs.append(msg)
    render_logs()


# ========================
# 環境
# ========================
st.subheader("1. 選擇環境")

env_option = st.radio(
    "環境",
    ["prod", "dev"],
    horizontal=True,
    index=0
)

memo.set_env(env_option)
st.caption(f"目前環境：{env_option}")


# ========================
# 帳密
# ========================
st.subheader("2. 帳號設定")

col1, col2 = st.columns(2)

with col1:
    st.markdown("#### 台北")
    accounts.ACCOUNTS["台北"]["email"] = st.text_input(
        "台北 Email",
        value=accounts.ACCOUNTS.get("台北", {}).get("email", ""),
    )
    accounts.ACCOUNTS["台北"]["password"] = st.text_input(
        "台北 Password",
        value=accounts.ACCOUNTS.get("台北", {}).get("password", ""),
        type="password",
    )

with col2:
    if "台中" not in accounts.ACCOUNTS:
        accounts.ACCOUNTS["台中"] = {"email": "", "password": "", "address_keywords": ["台中"]}
    st.markdown("#### 台中")
    accounts.ACCOUNTS["台中"]["email"] = st.text_input(
        "台中 Email",
        value=accounts.ACCOUNTS.get("台中", {}).get("email", ""),
    )
    accounts.ACCOUNTS["台中"]["password"] = st.text_input(
        "台中 Password",
        value=accounts.ACCOUNTS.get("台中", {}).get("password", ""),
        type="password",
    )


# ========================
# 模式
# ========================
st.subheader("3. 選擇模式")

mode = st.radio(
    "處理方式",
    [
        "Google Sheet列號模式",
        "直接輸入電話號碼模式",
        "搜尋條件模式",
    ],
    horizontal=True,
)

row_spec = ""
phone = ""
date_s = ""
limit = 10

if mode == "Google Sheet列號模式":
    row_spec = st.text_input("列號", "2,3,5-8")
    force = st.checkbox("強制重跑", value=False)
else:
    force = False

if mode == "直接輸入電話號碼模式":
    phone = st.text_input("電話號碼", "")
elif mode == "搜尋條件模式":
    col_a, col_b = st.columns(2)
    with col_a:
        date_s = st.text_input("訂購日期 YYYY/MM/DD", "2026/04/09")
    with col_b:
        limit = st.number_input("每次處理筆數", min_value=1, max_value=100, value=5)

    st.caption("條件：付款狀態 = 已付款，服務狀態 = 未處理")


# ========================
# 執行按鈕
# ========================
run = st.button("🚀 執行", use_container_width=True)

st.subheader("4. 執行過程")
render_logs()

st.subheader("5. 執行結果")

m1, m2, m3, m4 = st.columns(4)
last = st.session_state.last_result or {"processed": 0, "success": 0, "failed": 0, "skipped": 0, "errors": []}

m1.metric("執行筆數", last.get("processed", 0))
m2.metric("成功", last.get("success", 0))
m3.metric("失敗", last.get("failed", 0))
m4.metric("略過", last.get("skipped", 0))

if run:
    st.session_state.logs = []
    st.session_state.last_result = None
    ui_log("準備執行中...")

    try:
        if mode == "Google Sheet列號模式":
            result = memo.main(
                row_spec=row_spec,
                force=force,
                ui_logger=ui_log,
            )
        elif mode == "直接輸入電話號碼模式":
            result = memo.main_by_phone(
                phone=phone,
                ui_logger=ui_log,
            )
        else:
            result = memo.main_by_conditions(
                date_s=date_s,
                limit=int(limit),
                ui_logger=ui_log,
            )

        st.session_state.last_result = result

        m1.metric("執行筆數", result.get("processed", 0))
        m2.metric("成功", result.get("success", 0))
        m3.metric("失敗", result.get("failed", 0))
        m4.metric("略過", result.get("skipped", 0))

        st.success("執行完成")

    except Exception as e:
        ui_log(f"執行失敗：{e}")
        st.error(f"執行失敗：{e}")

final_result = st.session_state.last_result or {}
errors = final_result.get("errors", [])

if errors:
    st.subheader("6. 失敗明細")
    for err in errors:
        st.error(err)
