# -*- coding: utf-8 -*-
import streamlit as st
import memo

st.set_page_config(page_title="Memo 自動回填系統", layout="wide")

st.markdown("""
<style>
.block-container {
    padding-top: 0.35rem !important;
    padding-bottom: 0.5rem !important;
    max-width: 1180px !important;
}

h1 {
    font-size: 20px !important;
    margin-bottom: 6px !important;
}

[data-testid="stTextInput"],
[data-testid="stNumberInput"],
[data-testid="stDateInput"],
[data-testid="stSelectbox"] {
    margin-bottom: 0 !important;
}

[data-testid="stButton"] > button {
    background: #111 !important;
    color: white !important;
    border-radius: 8px !important;
    padding: 8px 0 !important;
}

[data-testid="stCode"] {
    font-size: 12px !important;
    max-height: 360px;
    overflow-y: auto;
}

[data-testid="stMetric"] {
    background: #f8f9fb !important;
    border: 1px solid #eaecf0 !important;
    border-radius: 9px !important;
    padding: 12px 14px !important;
}
</style>
""", unsafe_allow_html=True)

DEFAULT_RESULT = {
    "processed": 0,
    "success": 0,
    "failed": 0,
    "skipped": 0,
    "updated_orders": 0,
    "errors": [],
}

if "logs" not in st.session_state:
    st.session_state.logs = []
if "result" not in st.session_state:
    st.session_state.result = None


def normalize_result(r):
    base = DEFAULT_RESULT.copy()
    if isinstance(r, dict):
        base.update(r)
    if not isinstance(base.get("errors"), list):
        base["errors"] = []
    return base


def reset_state():
    st.session_state.logs = []
    st.session_state.result = None


def ui_log(msg):
    st.session_state.logs.append(str(msg))
    log_box.code("\n".join(st.session_state.logs[-3000:]))


st.title("📋 Memo 自動回填系統")

env_col1, env_col2 = st.columns([1, 8])
with env_col1:
    env_option = st.selectbox("", ["prod", "dev"], index=0, label_visibility="collapsed")

memo.set_env(env_option)

st.markdown("---")

c1, c2 = st.columns(2)
with c1:
    email = st.text_input("Email")
with c2:
    password = st.text_input("Password", type="password")

memo.set_runtime_credentials(email, password)

st.markdown("---")

mode = st.radio("", ["By Google Sheet", "By 電話", "By 搜尋條件"], horizontal=True)

row_spec = ""
phone = ""
date_mode = "服務日期"
purchase_status_name = "已付款"
limit = 5
start_date = None
end_date = None
force = False

if mode == "By Google Sheet":
    c1, c2 = st.columns([4, 1])
    with c1:
        row_spec = st.text_input("列號（例：2,3,5-8）")
    with c2:
        force = st.checkbox("強制重跑", value=False)

elif mode == "By 電話":
    phone = st.text_input("電話")

else:
    c1, c2, c3 = st.columns([1.2, 1.2, 1])
    with c1:
        date_mode = st.selectbox("日期條件", ["服務日期", "購買日期"])
    with c2:
        purchase_status_name = st.selectbox("付款狀態", ["未付款", "已付款"], index=1)
    with c3:
        limit = st.number_input("處理筆數", min_value=1, value=5)

    c4, c5 = st.columns(2)
    with c4:
        start_date = st.date_input("開始日期", value=None)
    with c5:
        end_date = st.date_input("結束日期", value=None)

run = st.button("🚀 執行", use_container_width=True)

st.markdown("---")
st.subheader("執行過程")
log_box = st.empty()
if st.session_state.logs:
    log_box.code("\n".join(st.session_state.logs[-3000:]))
else:
    log_box.code("尚未執行")

st.markdown("---")
result_container = st.container()


def render_result(result):
    r = normalize_result(result)
    with result_container:
        st.subheader("執行結果")
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("執行筆數", r["processed"])
        c2.metric("成功", r["success"])
        c3.metric("失敗", r["failed"])
        c4.metric("略過", r["skipped"])
        c5.metric("回寫筆數", r["updated_orders"])

        if r["errors"]:
            with st.expander(f"錯誤明細（{len(r['errors'])}）", expanded=True):
                for idx, err in enumerate(r["errors"], start=1):
                    st.markdown(f"**{idx}.** {err}")
        else:
            st.success("本次沒有錯誤")


if st.session_state.result:
    render_result(st.session_state.result)

if run:
    reset_state()
    log_box.code("尚未執行")
    result_container.empty()

    if not email or not password:
        st.session_state.result = {**DEFAULT_RESULT, "failed": 1, "errors": ["請先輸入 Email / Password"]}
        render_result(st.session_state.result)
        st.stop()

    try:
        ui_log("===== 開始執行 =====")

        if mode == "By Google Sheet":
            result = memo.main(row_spec=row_spec, force=force, ui_logger=ui_log)

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
        render_result(result)

    except Exception as e:
        ui_log(f"❌ 執行錯誤: {e}")
        st.session_state.result = {**DEFAULT_RESULT, "failed": 1, "errors": [str(e)]}
        render_result(st.session_state.result)
