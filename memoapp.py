import streamlit as st
import memo
import accounts

st.set_page_config(page_title="Memo System", layout="wide")

st.title("📋 Memo 自動補備註系統")

env_options = {
    "正式機 prod": "prod",
    "測試機 dev": "dev",
}

selected_env = st.radio("環境", list(env_options.keys()), horizontal=True)
memo.set_env(env_options[selected_env])

# 帳密
st.subheader("登入帳密")

col1, col2 = st.columns(2)

with col1:
    accounts.ACCOUNTS["台北"]["email"] = st.text_input("台北 email")
    accounts.ACCOUNTS["台北"]["password"] = st.text_input("台北 password", type="password")

with col2:
    accounts.ACCOUNTS["台中"]["email"] = st.text_input("台中 email")
    accounts.ACCOUNTS["台中"]["password"] = st.text_input("台中 password", type="password")

# 列號
row_spec = st.text_input("處理列", value="2,3,5-8")
force = st.checkbox("強制重跑")

log_box = st.empty()

def ui_logger(msg):
    st.session_state.logs.append(msg)
    log_box.code("\n".join(st.session_state.logs[-500:]))

if "logs" not in st.session_state:
    st.session_state.logs = []

if st.button("🚀 執行"):
    st.session_state.logs = []
    result = memo.main(row_spec=row_spec, force=force, ui_logger=ui_logger)

    st.success("完成")
    st.write(result)
