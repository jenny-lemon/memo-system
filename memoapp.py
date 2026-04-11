import streamlit as st
import requests
import memo

st.set_page_config(layout="wide")

st.title("📋 Memo 自動回填系統")

# ========================
# 初始化
# ========================
if "logs" not in st.session_state:
    st.session_state.logs = []

def log(msg):
    st.session_state.logs.append(msg)
    st.code("\n".join(st.session_state.logs))

# ========================
# UI
# ========================
email = st.text_input("Email")
password = st.text_input("Password", type="password")

mode = st.radio("模式", ["單筆訂單", "電話批次"])

order_no = ""
phone = ""

if mode == "單筆訂單":
    order_no = st.text_input("訂單編號")

if mode == "電話批次":
    phone = st.text_input("電話")

run = st.button("🚀 執行")

# ========================
# 執行
# ========================
if run:

    st.session_state.logs = []

    session = requests.Session()

    try:
        memo.login(session, email, password)
        log("[登入] 成功")

        if mode == "單筆訂單":

            count, ok = memo.process_order(session, order_no, log)

            if ok:
                st.success(f"✅ 成功：回寫 {count} 筆")
            else:
                st.error("❌ 失敗")

        elif mode == "電話批次":

            items = memo.parse_list(
                session.get(f"{memo.BASE_URL}/purchase?phone={phone}&purchase_status=1").text
            )

            total = 0
            success = 0

            for i in items:
                c, ok = memo.process_order(session, i["order_no"], log)
                total += c
                if ok:
                    success += 1

            st.success(f"完成：成功 {success} / 共 {len(items)} 筆")

    except Exception as e:
        st.error(f"❌ 錯誤：{e}")
