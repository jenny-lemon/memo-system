# -*- coding: utf-8 -*-
import streamlit as st
import memo
from datetime import date

st.set_page_config(page_title="Memo 自動回填系統", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@300;400;500;600&family=DM+Sans:wght@300;400;500;600&display=swap');

html, body, [class*="css"] {
    font-family: 'DM Sans', 'Noto Sans TC', sans-serif;
}

/* 再縮上方與區塊間距 */
.block-container {
    padding-top: 0.8rem !important;
    padding-bottom: 0.8rem !important;
    max-width: 1200px !important;
}
[data-testid="stVerticalBlock"] > div {
    gap: 0.2rem !important;
}

h1 {
    font-size: 18px !important;
    font-weight: 600 !important;
    margin: 0 0 8px 0 !important;
}
h3 {
    font-size: 11px !important;
    font-weight: 500 !important;
    color: #8a8f98 !important;
    letter-spacing: 0.08em !important;
    text-transform: uppercase !important;
    margin: 8px 0 2px 0 !important;
}
hr {
    border-color: rgba(0,0,0,0.08) !important;
    margin: 6px 0 !important;
}

[data-testid="stRadio"] {
    margin-bottom: 0 !important;
}
[data-testid="stRadio"] > div {
    gap: 0.4rem !important;
    margin-top: 0 !important;
}
[data-testid="stRadio"] label {
    font-size: 13px !important;
}

[data-testid="stTextInput"],
[data-testid="stNumberInput"],
[data-testid="stDateInput"],
[data-testid="stSelectbox"] {
    margin-bottom: 0 !important;
}
[data-testid="stTextInput"] label,
[data-testid="stNumberInput"] label,
[data-testid="stDateInput"] label,
[data-testid="stSelectbox"] label {
    font-size: 12px !important;
    margin-bottom: 2px !important;
    color: #666 !important;
}
[data-testid="stTextInput"] input,
[data-testid="stNumberInput"] input,
[data-testid="stDateInput"] input {
    font-size: 13px !important;
    padding: 6px 9px !important;
    border-radius: 8px !important;
}

[data-testid="stCheckbox"] {
    margin-top: 0.25rem !important;
    margin-bottom: 0 !important;
}

[data-testid="stButton"] > button {
    background: #16181d !important;
    color: #fff !important;
    border: none !important;
    border-radius: 9px !important;
    font-size: 14px !important;
    font-weight: 500 !important;
    padding: 9px 0 !important;
}
[data-testid="stButton"] > button:hover {
    background: #2a2d33 !important;
}

[data-testid="stCode"] {
    font-size: 12px !important;
    border-radius: 10px !important;
    max-height: 320px;
    overflow-y: auto;
    margin-top: 4px !important;
}

[data-testid="stMetric"] {
    background: #f8f9fb !important;
    border: 1px solid #e8ebf0 !important;
    border-radius: 10px !important;
    padding: 10px 12px !important;
}
[data-testid="stMetricLabel"] {
    font-size: 12px !important;
    color: #8a8f98 !important;
}
[data-testid="stMetricValue"] {
    font-size: 22px !important;
    font-weight: 600 !important;
}

[data-testid="stAlert"] {
    border-radius: 10px !important;
    margin-top: 8px !important;
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
if "last_result" not in st.session_state:
    st.session_state.last_result = None
if "ran" not in st.session_state:
    st.session_state.ran = False


def normalize_result(r):
    base = DEFAULT_RESULT.copy()
    if isinstance(r, dict):
        base.update(r)
    if not isinstance(base.get("errors"), list):
        base["errors"] = []
    return base


def reset_run_state():
    st.session_state.logs = []
    st.session_state.last_result = None
    st.session_state.ran = False


st.title("📋 Memo 自動回填系統")
st.markdown("<hr>", unsafe_allow_html=True)

st.subheader("環境")
env_option = st.radio("環境", ["prod", "dev"], horizontal=True, index=0, label_visibility="collapsed")
memo.set_env(env_option)
st.caption(f"目前環境：{env_option} | {getattr(memo, 'BASE_URL', '')}")

st.markdown("<hr>", unsafe_allow_html=True)

st.subheader("登入帳密")
c1, c2 = st.columns(2)
with c1:
    email = st.text_input("Email", key="login_email")
with c2:
    password = st.text_input("Password", type="password", key="login_password")

memo.set_runtime_credentials(email, password)

st.markdown("<hr>", unsafe_allow_html=True)

st.subheader("處理模式")
mode = st.radio(
    "mode",
    ["By Google Sheet 列號", "By 電話", "By 搜尋條件"],
    horizontal=True,
    label_visibility="collapsed",
)

row_spec = ""
phone = ""
force = False

date_field = "服務日期"
date_start_text = ""
date_end_text = ""
date_start_picker = None
date_end_picker = None
purchase_status_name = "已付款"
limit = 5

if mode == "By Google Sheet 列號":
    c1, c2 = st.columns([4, 1])
    with c1:
        row_spec = st.text_input("列號（例：2,3,5-8）", "2,3,5-8")
    with c2:
        st.markdown("<div style='height:22px'></div>", unsafe_allow_html=True)
        force = st.checkbox("強制重跑", value=False)

elif mode == "By 電話":
    phone = st.text_input("電話號碼")

else:
    c1, c2, c3 = st.columns([1.2, 1.2, 1])
    with c1:
        date_field = st.selectbox("日期條件", ["服務日期", "購買日期"], index=0)
    with c2:
        purchase_status_name = st.selectbox("付款狀態", ["已付款", "未付款"], index=0)
    with c3:
        limit = st.number_input("處理筆數", min_value=1, max_value=100, value=5)

    c4, c5 = st.columns(2)
    with c4:
        date_start_text = st.text_input("起日（可手打 YYYY/MM/DD）", "")
        date_start_picker = st.date_input("或用日曆選起日", value=None)
    with c5:
        date_end_text = st.text_input("迄日（可手打 YYYY/MM/DD）", "")
        date_end_picker = st.date_input("或用日曆選迄日", value=None)

    st.caption("搜尋條件：可選服務日期或購買日期區間，再搭配付款狀態與處理筆數。")

st.markdown("<hr>", unsafe_allow_html=True)

run = st.button("🚀 執行", use_container_width=True)

st.subheader("執行過程")
log_placeholder = st.empty()


def render_logs():
    text = "\n".join(st.session_state.logs[-3000:]) if st.session_state.logs else "尚未執行"
    log_placeholder.code(text)


def ui_log(msg: str):
    st.session_state.logs.append(str(msg))
    log_placeholder.code("\n".join(st.session_state.logs[-3000:]))


render_logs()

result_area = st.empty()


def render_result(result):
    r = normalize_result(result)

    with result_area.container():
        st.markdown("<hr>", unsafe_allow_html=True)
        st.subheader("執行結果")

        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("執行筆數", r["processed"])
        m2.metric("成功", r["success"])
        m3.metric("失敗", r["failed"])
        m4.metric("略過", r["skipped"])
        m5.metric("回寫筆數", r["updated_orders"])

        if r["failed"] == 0 and r["processed"] > 0:
            st.success(f"✅ 全部執行完成，共處理 {r['processed']} 筆，成功 {r['success']} 筆。")
        elif r["failed"] > 0:
            st.warning(f"⚠️ 執行完成，但有 {r['failed']} 筆失敗，請查看下方錯誤明細。")
        else:
            st.info("執行完成，無資料被處理。")

        if r["errors"]:
            st.subheader("錯誤明細")
            for err in r["errors"]:
                st.error(err)


if st.session_state.ran and st.session_state.last_result:
    render_result(st.session_state.last_result)

if run:
    reset_run_state()
    result_area.empty()
    render_logs()

    if not email or not password:
        result = {**DEFAULT_RESULT, "failed": 1, "errors": ["請先輸入 Email / Password"]}
        st.session_state.last_result = result
        st.session_state.ran = True
        render_result(result)
        st.stop()

    try:
        ui_log("===== 開始執行 =====")

        if mode == "By Google Sheet 列號":
            result = memo.main(
                row_spec=row_spec,
                force=force,
                ui_logger=ui_log,
            )

        elif mode == "By 電話":
            result = memo.main_by_phone(
                phone=phone,
                ui_logger=ui_log,
            )

        else:
            start_date = date_start_text.strip() or (date_start_picker.strftime("%Y/%m/%d") if date_start_picker else "")
            end_date = date_end_text.strip() or (date_end_picker.strftime("%Y/%m/%d") if date_end_picker else "")

            result = memo.main_by_conditions(
                date_mode=date_field,
                date_start=start_date,
                date_end=end_date,
                purchase_status_name=purchase_status_name,
                limit=int(limit),
                ui_logger=ui_log,
            )

        ui_log("===== 執行完成 =====")
        st.session_state.last_result = result
        st.session_state.ran = True
        render_result(result)

    except Exception as e:
        err = f"執行失敗：{e}"
        ui_log(err)
        result = {**DEFAULT_RESULT, "failed": 1, "errors": [str(e)]}
        st.session_state.last_result = result
        st.session_state.ran = True
        render_result(result)
