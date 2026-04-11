# -*- coding: utf-8 -*-
import streamlit as st
import memo

st.set_page_config(page_title="Memo 自動回填系統", layout="wide")

# ================== STYLE ==================
st.markdown("""
<style>
.block-container {
    padding-top: 0.35rem !important;
    padding-bottom: 0.5rem !important;
    max-width: 1100px !important;
}

/* 標題 */
h1 {
    font-size: 18px !important;
    margin-bottom: 6px !important;
}

/* input */
[data-testid="stTextInput"],
[data-testid="stNumberInput"],
[data-testid="stDateInput"],
[data-testid="stSelectbox"] {
    margin-bottom: 0 !important;
}

[data-testid="stTextInput"] input,
[data-testid="stNumberInput"] input,
[data-testid="stDateInput"] input {
    font-size: 13px !important;
    padding: 6px 8px !important;
}

/* button */
[data-testid="stButton"] > button {
    background: #111 !important;
    color: white !important;
    border-radius: 6px !important;
    padding: 8px 0 !important;
}

/* log */
[data-testid="stCode"] {
    font-size: 12px !important;
    max-height: 280px;
}
</style>
""", unsafe_allow_html=True)

# ================== STATE ==================
if "logs" not in st.session_state:
    st.session_state.logs = []
if "result" not in st.session_state:
    st.session_state.result = None

def ui_log(msg):
    st.session_state.logs.append(str(msg))

# ================== TITLE ==================
st.title("📋 Memo 自動回填系統")

# ================== ENV（縮小）==================
col_env1, col_env2 = st.columns([1, 6])
with col_env1:
    env_option = st.selectbox(
        "",
        ["prod", "dev"],
        index=0,
        label_visibility="collapsed"
    )

memo.set_env(env_option)

# ================== LOGIN ==================
st.markdown("---")

col1, col2 = st.columns(2)
with col1:
    email = st.text_input("Email")
with col2:
    password = st.text_input("Password", type="password")

memo.set_runtime_credentials(email, password)

# ================== MODE ==================
st.markdown("---")

mode = st.radio(
    "",
    ["By Google Sheet", "By 電話", "By 搜尋條件"],
    horizontal=True
)

row_spec = ""
phone = ""
limit = 5
date_mode = "服務日期"
purchase_status_name = "已付款"
start_date = None
end_date = None

# ================== INPUT ==================
if mode == "By Google Sheet":
    row_spec = st.text_input("列號（例：2,3,5-8）")

elif mode == "By 電話":
    phone = st.text_input("電話")

else:
    col1, col2, col3 = st.columns([1.2, 1.2, 1])
    with col1:
        date_mode = st.selectbox("日期條件", ["服務日期", "購買日期"])
    with col2:
        purchase_status_name = st.selectbox("付款狀態", ["未付款", "已付款"])
    with col3:
        limit = st.number_input("處理筆數", min_value=1, value=5)

    col4, col5 = st.columns(2)
    with col4:
        start_date = st.date_input("開始日期", value=None)
    with col5:
        end_date = st.date_input("結束日期", value=None)

# ================== RUN ==================
st.markdown("---")
run = st.button("🚀 執行", use_container_width=True)

# ================== LOG ==================
st.subheader("執行過程")

log_box = st.empty()

def render_log():
    txt = "\n".join(st.session_state.logs[-2000:]) if st.session_state.logs else "尚未執行"
    log_box.code(txt)

render_log()

# ================== RESULT ==================
result_box = st.empty()

def render_result(r):
    if not r:
        return

    with result_box.container():
        st.markdown("---")
        st.subheader("執行結果")

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("執行筆數", r.get("processed", 0))
        c2.metric("成功", r.get("success", 0))
        c3.metric("失敗", r.get("failed", 0))
        c4.metric("略過", r.get("skipped", 0))
        c5.metric("回寫筆數", r.get("updated_orders", 0))

        if r.get("errors"):
            st.subheader("錯誤明細")
            for e in r["errors"]:
                st.error(e)

# ================== EXECUTE ==================
if run:
    st.session_state.logs = []
    result_box.empty()

    if not email or not password:
        ui_log("❌ 請輸入帳密")
        render_log()
        st.stop()

    try:
        ui_log("===== 開始執行 =====")

        if mode == "By Google Sheet":
            result = memo.main(row_spec=row_spec, ui_logger=ui_log)

        elif mode == "By 電話":
            result = memo.main_by_phone(phone=phone, ui_logger=ui_log)

        else:
            start_text = start_date.strftime("%Y/%m/%d") if start_date else ""
            end_text = end_date.strftime("%Y/%m/%d") if end_date else ""

            result = memo.main_by_conditions(
                date_mode=date_mode,
                date_start=start_text,
                date_end=end_text,
                purchase_status_name=purchase_status_name,
                limit=int(limit),
                ui_logger=ui_log,
            )

        ui_log("===== 執行完成 =====")

        st.session_state.result = result

        render_log()
        render_result(result)

    except Exception as e:
        ui_log(f"❌ 執行錯誤: {e}")
        render_log()
