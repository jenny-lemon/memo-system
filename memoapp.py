# -*- coding: utf-8 -*-
import streamlit as st
import memo

st.set_page_config(page_title="Memo 自動回填系統", layout="wide")

st.title("📋 Memo 自動回填系統")

# ========================
# 初始化
# ========================
if "logs" not in st.session_state:
    st.session_state.logs = []

if "result" not in st.session_state:
    st.session_state.result = {
        "processed": 0,
        "success": 0,
        "failed": 0,
        "skipped": 0,
        "updated_orders": 0,
        "errors": []
    }

def reset():
    st.session_state.logs = []
    st.session_state.result = {
        "processed": 0,
        "success": 0,
        "failed": 0,
        "skipped": 0,
        "updated_orders": 0,
        "errors": []
    }

def ui_log(msg):
    st.session_state.logs.append(str(msg))
    log_box.code("\n".join(st.session_state.logs[-2000:]))

# ========================
# 帳密
# ========================
st.subheader("🔐 登入")

email = st.text_input("Email")
password = st.text_input("Password", type="password")

memo.set_runtime_credentials(email, password)

# ========================
# 模式
# ========================
st.subheader("⚙️ 執行模式")

mode = st.radio(
    "選擇模式",
    [
        "By Google Sheet 列號",
        "By 電話",
        "By 搜尋條件"
    ],
    horizontal=True
)

row_spec = ""
phone = ""
date_s = ""
limit = 5
force = False

if mode == "By Google Sheet 列號":
    row_spec = st.text_input("列號", "2,3,5-10")
    force = st.checkbox("強制重跑")

elif mode == "By 電話":
    phone = st.text_input("電話")

elif mode == "By 搜尋條件":
    c1, c2 = st.columns(2)
    with c1:
        date_s = st.text_input("訂購日期 YYYY/MM/DD")
    with c2:
        limit = st.number_input("每次處理筆數", 1, 50, 5)

    st.caption("條件：付款=已付款 + 服務狀態=未處理")

# ========================
# 執行按鈕
# ========================
run = st.button("🚀 執行", use_container_width=True)

# ========================
# LOG顯示
# ========================
st.subheader("📜 執行過程")
log_box = st.empty()

if not st.session_state.logs:
    log_box.code("尚未執行")

# ========================
# 結果顯示
# ========================
st.subheader("📊 執行結果")

def render_result():
    r = st.session_state.result

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("執行筆數", r["processed"])
    c2.metric("成功", r["success"])
    c3.metric("失敗", r["failed"])
    c4.metric("略過", r["skipped"])
    c5.metric("回寫筆數", r["updated_orders"])

    if r["errors"]:
        st.subheader("❌ 錯誤")
        for e in r["errors"]:
            st.error(e)

render_result()

# ========================
# 執行
# ========================
if run:

    reset()

    if not email or not password:
        st.error("請輸入帳密")
        st.stop()

    try:
        ui_log("===== 開始執行 =====")

        if mode == "By Google Sheet 列號":

            result = memo.main(
                row_spec=row_spec,
                force=force,
                ui_logger=ui_log
            )

        elif mode == "By 電話":

            result = memo.main_by_phone(
                phone=phone,
                ui_logger=ui_log
            )

        elif mode == "By 搜尋條件":

            result = memo.main_by_conditions(
                date_s=date_s,
                limit=int(limit),
                ui_logger=ui_log
            )

        st.session_state.result = result

        ui_log("===== 執行完成 =====")

    except Exception as e:
        ui_log(f"❌ 系統錯誤: {e}")
        st.session_state.result["failed"] += 1
        st.session_state.result["errors"].append(str(e))

    render_result()
