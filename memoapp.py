# -*- coding: utf-8 -*-
import streamlit as st
import memo

st.set_page_config(page_title="Memo 自動回填系統", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@300;400;500;600&family=DM+Sans:opsz,wght@9..40,300;9..40,400;9..40,500;9..40,600&display=swap');

html, body, [class*="css"] {
    font-family: 'DM Sans', 'Noto Sans TC', sans-serif;
}
.block-container {
    padding-top: 1.4rem !important;
    padding-bottom: 1rem !important;
    max-width: 1100px !important;
}

/* 標題 */
h1 {
    font-size: 19px !important;
    font-weight: 600 !important;
    margin: 0 0 14px 0 !important;
    display: flex; align-items: center; gap: 8px;
}

/* section label */
.sec-label {
    font-size: 11px;
    font-weight: 600;
    color: #aaa;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    margin: 0 0 6px 0;
}

/* Input / Select */
[data-testid="stTextInput"] label,
[data-testid="stNumberInput"] label,
[data-testid="stDateInput"] label,
[data-testid="stSelectbox"] label {
    font-size: 11.5px !important;
    color: #777 !important;
    font-weight: 500 !important;
    margin-bottom: 2px !important;
}
[data-testid="stTextInput"] input,
[data-testid="stNumberInput"] input {
    font-size: 13.5px !important;
    padding: 7px 10px !important;
    border-radius: 7px !important;
    border-color: #dde0e5 !important;
    background: #fff !important;
}
[data-testid="stTextInput"] input:focus,
[data-testid="stNumberInput"] input:focus {
    border-color: #111 !important;
    box-shadow: 0 0 0 2px rgba(0,0,0,0.07) !important;
}

/* Selectbox */
[data-testid="stSelectbox"] > div > div {
    border-radius: 7px !important;
    border-color: #dde0e5 !important;
    font-size: 13.5px !important;
    background: #fff !important;
}

/* Radio */
[data-testid="stRadio"] > div { gap: 6px !important; }
[data-testid="stRadio"] label span { font-size: 13px !important; }

/* Checkbox */
[data-testid="stCheckbox"] label span { font-size: 13px !important; color: #555 !important; }

/* 執行按鈕 */
[data-testid="stButton"] > button {
    background: #111 !important;
    color: #fff !important;
    border: none !important;
    border-radius: 8px !important;
    font-size: 14px !important;
    font-weight: 600 !important;
    padding: 9px 0 !important;
    letter-spacing: 0.02em !important;
    transition: background 0.15s !important;
}
[data-testid="stButton"] > button:hover { background: #333 !important; }
[data-testid="stButton"] > button:disabled {
    background: #ccc !important;
    cursor: not-allowed !important;
}

/* Expander */
[data-testid="stExpander"] {
    border: 1px solid #e4e6ea !important;
    border-radius: 10px !important;
    background: #fafafa !important;
    overflow: hidden !important;
}
[data-testid="stExpander"] summary {
    font-size: 12px !important;
    font-weight: 600 !important;
    color: #666 !important;
    letter-spacing: 0.04em !important;
    padding: 10px 14px !important;
}

/* Log */
[data-testid="stCode"] {
    font-size: 11.5px !important;
    border-radius: 0 0 10px 10px !important;
    max-height: 280px;
    overflow-y: auto;
    background: #13131f !important;
    margin: 0 !important;
}

/* Metric */
[data-testid="stMetric"] {
    background: #fff !important;
    border: 1px solid #e4e6ea !important;
    border-radius: 10px !important;
    padding: 14px 16px 12px !important;
    text-align: center !important;
}
[data-testid="stMetricLabel"] {
    font-size: 11px !important;
    color: #999 !important;
    font-weight: 600 !important;
    text-transform: uppercase !important;
    letter-spacing: 0.06em !important;
}
[data-testid="stMetricValue"] {
    font-size: 30px !important;
    font-weight: 700 !important;
    color: #111 !important;
}

/* divider */
hr { border-color: #ebebeb !important; margin: 12px 0 !important; }

/* Alert */
[data-testid="stAlert"] {
    border-radius: 8px !important;
    font-size: 13px !important;
    margin-top: 10px !important;
}

/* Caption */
[data-testid="stCaptionContainer"] {
    font-size: 11.5px !important;
    color: #bbb !important;
    margin-top: 2px !important;
}

/* date input */
[data-testid="stDateInput"] input {
    font-size: 13px !important;
    border-radius: 7px !important;
    border-color: #dde0e5 !important;
}
</style>
""", unsafe_allow_html=True)

# ── Session state ─────────────────────────────────────────
DEFAULT_RESULT = {
    "processed": 0, "success": 0, "failed": 0,
    "skipped": 0, "updated_orders": 0, "errors": [],
}
if "logs"       not in st.session_state: st.session_state.logs = []
if "result"     not in st.session_state: st.session_state.result = None
if "is_running" not in st.session_state: st.session_state.is_running = False

def normalize_result(r):
    base = DEFAULT_RESULT.copy()
    if isinstance(r, dict): base.update(r)
    if not isinstance(base.get("errors"), list): base["errors"] = []
    return base

def ui_log(msg):
    st.session_state.logs.append(str(msg))
    log_box.code("\n".join(st.session_state.logs[-3000:]))

def sec(title):
    st.markdown(f'<p class="sec-label">{title}</p>', unsafe_allow_html=True)

def render_result(result):
    r = normalize_result(result)
    with result_container:
        st.markdown("<hr>", unsafe_allow_html=True)
        sec("執行結果")
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("執行筆數", r["processed"])
        c2.metric("成功",     r["success"])
        c3.metric("失敗",     r["failed"])
        c4.metric("略過",     r["skipped"])
        c5.metric("回寫筆數", r["updated_orders"])

        if r["errors"]:
            with st.expander(f"⚠️  錯誤明細（{len(r['errors'])} 筆）", expanded=True):
                for i, err in enumerate(r["errors"], 1):
                    st.markdown(f"**{i}.** {err}")
        elif r["processed"] > 0:
            st.success(f"✅ 全部完成，共處理 **{r['processed']}** 筆，成功 **{r['success']}** 筆。")
        else:
            st.info("執行完成，無資料被處理。")

# ── 標題 ─────────────────────────────────────────────────
st.title("📋 Memo 自動回填系統")

# ── 帳密 + 環境（同列）────────────────────────────────────
col_e, col_p, col_env = st.columns([3.2, 3.2, 1.2])
with col_e:
    email    = st.text_input("Email")
with col_p:
    password = st.text_input("Password", type="password")
with col_env:
    env_option = st.selectbox("環境", ["prod", "dev"], index=0)

memo.set_env(env_option)
memo.set_runtime_credentials(email, password)

st.markdown("<hr>", unsafe_allow_html=True)

# ── 處理模式 ──────────────────────────────────────────────
sec("處理模式")
mode = st.radio(
    "", ["By Google Sheet", "By 電話", "By 搜尋條件"],
    horizontal=True, label_visibility="collapsed"
)

row_spec = phone = ""
date_mode = "服務日期"
purchase_status_name = "已付款"
limit = 5
start_date = end_date = None
force = False

if mode == "By Google Sheet":
    c1, c2 = st.columns([5, 1])
    with c1:
        row_spec = st.text_input("列號（例：2,3,5-8）")
    with c2:
        st.markdown("<div style='height:26px'></div>", unsafe_allow_html=True)
        force = st.checkbox("強制重跑")

elif mode == "By 電話":
    phone = st.text_input("電話號碼")

else:
    c1, c2, c3 = st.columns([1.5, 1.5, 1])
    with c1: date_mode             = st.selectbox("日期條件", ["服務日期", "購買日期"])
    with c2: purchase_status_name  = st.selectbox("付款狀態", ["未付款", "已付款"], index=1)
    with c3: limit                 = st.number_input("處理筆數", min_value=1, value=5)
    c4, c5 = st.columns(2)
    with c4: start_date = st.date_input("開始日期", value=None)
    with c5: end_date   = st.date_input("結束日期",  value=None)

st.markdown("<hr>", unsafe_allow_html=True)

# ── 執行按鈕 ──────────────────────────────────────────────
run = st.button("🚀  執行", use_container_width=True, disabled=st.session_state.is_running)

# ── 執行過程 ──────────────────────────────────────────────
with st.expander("📄  執行過程", expanded=True):
    log_box = st.empty()
    log_box.code("\n".join(st.session_state.logs[-3000:]) if st.session_state.logs else "尚未執行")

# ── 執行結果 ──────────────────────────────────────────────
result_container = st.container()
if st.session_state.result is not None:
    render_result(st.session_state.result)

# ── 執行邏輯 ─────────────────────────────────────────────
if run:
    st.session_state.logs = []
    st.session_state.result = None
    log_box.code("尚未執行")

    if not email or not password:
        st.session_state.result = {**DEFAULT_RESULT, "failed": 1, "errors": ["請先輸入 Email / Password"]}
        render_result(st.session_state.result)
        st.stop()

    try:
        st.session_state.is_running = True
        ui_log("===== 開始執行 =====")

        with st.spinner("執行中，請稍候…"):
            if mode == "By Google Sheet":
                result = memo.main(row_spec=row_spec, force=force, ui_logger=ui_log)
            elif mode == "By 電話":
                result = memo.main_by_phone(phone=phone, ui_logger=ui_log)
            else:
                start_text = start_date.strftime("%Y/%m/%d") if start_date else ""
                end_text   = end_date.strftime("%Y/%m/%d")   if end_date   else ""
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
        ui_log(f"❌ 執行錯誤：{e}")
        st.session_state.result = {**DEFAULT_RESULT, "failed": 1, "errors": [str(e)]}
        render_result(st.session_state.result)

    finally:
        st.session_state.is_running = False
