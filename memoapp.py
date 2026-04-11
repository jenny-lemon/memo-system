# -*- coding: utf-8 -*-
import streamlit as st
import memo

st.set_page_config(page_title="Memo 自動回填系統", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@300;400;500;600&family=DM+Sans:opsz,wght@9..40,300;9..40,400;9..40,500;9..40,600&display=swap');

html, body, [class*="css"] { font-family: 'DM Sans', 'Noto Sans TC', sans-serif; }

/* 大幅縮減 block 間距 */
.block-container { padding-top: 1.5rem !important; padding-bottom: 1rem !important; }
[data-testid="stVerticalBlock"] > [data-testid="stVerticalBlockBorderWrapper"],
[data-testid="stVerticalBlock"] > div { gap: 0 !important; }

/* 壓縮各元件的 margin */
h1 { font-size: 20px !important; font-weight: 600 !important; margin: 0 0 12px 0 !important; }
h3 { font-size: 11.5px !important; font-weight: 500 !important; color: #888 !important;
     letter-spacing: 0.08em !important; text-transform: uppercase !important;
     margin: 14px 0 4px 0 !important; }
hr { border-color: rgba(0,0,0,0.08) !important; margin: 10px 0 !important; }

/* Radio */
[data-testid="stRadio"] { margin-bottom: 2px !important; }
[data-testid="stRadio"] > label { margin-bottom: 2px !important; font-size: 13px !important; }
[data-testid="stRadio"] > div { gap: 6px !important; margin-top: 2px !important; }

/* Caption */
[data-testid="stCaptionContainer"] { margin-top: 2px !important; margin-bottom: 0 !important;
    font-size: 12px !important; color: #999 !important; }

/* Text input */
[data-testid="stTextInput"] { margin-bottom: 0 !important; }
[data-testid="stTextInput"] label { font-size: 12.5px !important; margin-bottom: 2px !important; color: #555 !important; }
[data-testid="stTextInput"] input { font-size: 13.5px !important; padding: 7px 10px !important; border-radius: 7px !important; }

/* Checkbox */
[data-testid="stCheckbox"] { margin-top: 6px !important; }
[data-testid="stCheckbox"] label { font-size: 13px !important; }

/* Number input */
[data-testid="stNumberInput"] label { font-size: 12.5px !important; color: #555 !important; }

/* 執行按鈕 — 黑色 */
[data-testid="stButton"] > button {
    background: #1a1a1a !important;
    border: none !important;
    color: #fff !important;
    font-weight: 500 !important;
    font-size: 14px !important;
    border-radius: 8px !important;
    padding: 9px 0 !important;
    letter-spacing: 0.02em !important;
    transition: background 0.18s !important;
    box-shadow: none !important;
}
[data-testid="stButton"] > button:hover { background: #333 !important; }

/* Log */
[data-testid="stCode"] {
    font-size: 12px !important; border-radius: 8px !important;
    max-height: 280px; overflow-y: auto;
    margin-top: 4px !important;
}

/* Metric 卡片 */
[data-testid="stMetric"] {
    background: #f8f9fb !important;
    border: 1px solid #eaecf0 !important;
    border-radius: 9px !important;
    padding: 12px 14px !important;
}
[data-testid="stMetricLabel"] { font-size: 12px !important; color: #888 !important; }
[data-testid="stMetricValue"] { font-size: 24px !important; font-weight: 600 !important; }

/* Alert */
[data-testid="stAlert"] { border-radius: 8px !important; font-size: 13px !important; margin-top: 8px !important; }
</style>
""", unsafe_allow_html=True)

# ── Session state ─────────────────────────────────────────
DEFAULT_RESULT = {
    "processed": 0, "success": 0, "failed": 0,
    "skipped": 0, "updated_orders": 0, "errors": [],
}
if "logs"        not in st.session_state: st.session_state.logs = []
if "last_result" not in st.session_state: st.session_state.last_result = None
if "ran"         not in st.session_state: st.session_state.ran = False

def normalize_result(r):
    base = DEFAULT_RESULT.copy()
    if isinstance(r, dict): base.update(r)
    if not isinstance(base.get("errors"), list): base["errors"] = []
    return base

# ── 標題 ─────────────────────────────────────────────────
st.title("📋 Memo 自動回填系統")
st.markdown("<hr>", unsafe_allow_html=True)

# ── 1. 環境 ───────────────────────────────────────────────
st.subheader("環境")
env_option = st.radio("環境", ["prod", "dev"], horizontal=True, index=0, label_visibility="collapsed")
if hasattr(memo, "set_env"): memo.set_env(env_option)
base_url = getattr(memo, "BASE_URL", "")
st.caption(f"目前環境：{env_option}{f'  |  {base_url}' if base_url else ''}")

st.markdown("<hr>", unsafe_allow_html=True)

# ── 2. 帳密 ───────────────────────────────────────────────
st.subheader("登入帳密")
col_e, col_p = st.columns(2)
with col_e: email    = st.text_input("Email", key="login_email")
with col_p: password = st.text_input("Password", type="password", key="login_password")

if hasattr(memo, "set_runtime_credentials"):
    memo.set_runtime_credentials(email, password)
else:
    setattr(memo, "RUNTIME_EMAIL", email)
    setattr(memo, "RUNTIME_PASSWORD", password)

st.markdown("<hr>", unsafe_allow_html=True)

# ── 3. 模式 ───────────────────────────────────────────────
st.subheader("處理模式")
mode = st.radio("mode", ["By Google Sheet 列號", "By 電話", "By 搜尋條件"],
                horizontal=True, label_visibility="collapsed")

row_spec = phone = date_s = ""
limit = 5; force = False

if mode == "By Google Sheet 列號":
    col_r, col_f = st.columns([3, 1])
    with col_r: row_spec = st.text_input("列號（例：2,3,5-8）", "2,3,5-8")
    with col_f:
        st.markdown("<div style='height:26px'></div>", unsafe_allow_html=True)
        force = st.checkbox("強制重跑", value=False)
elif mode == "By 電話":
    phone = st.text_input("電話號碼")
elif mode == "By 搜尋條件":
    c1, c2 = st.columns(2)
    with c1: date_s = st.text_input("訂購日期 YYYY/MM/DD")
    with c2: limit  = st.number_input("每次處理筆數", min_value=1, max_value=100, value=5)
    st.caption("條件：付款狀態 ＝ 已付款，服務狀態 ＝ 未處理")

st.markdown("<hr>", unsafe_allow_html=True)

# ── 執行按鈕 ──────────────────────────────────────────────
run = st.button("🚀  執行", use_container_width=True)

# ── 4. 執行過程 ───────────────────────────────────────────
st.markdown("<hr>", unsafe_allow_html=True)
st.subheader("執行過程")
log_placeholder = st.empty()

def render_logs():
    text = "\n".join(st.session_state.logs[-3000:]) if st.session_state.logs else "尚未執行"
    log_placeholder.code(text)

def ui_log(msg: str):
    st.session_state.logs.append(str(msg))
    log_placeholder.code("\n".join(st.session_state.logs[-3000:]))

render_logs()

# ── 5. 執行結果（只在執行後顯示一次） ─────────────────────
result_area = st.empty()

def render_result(result):
    r = normalize_result(result)
    with result_area.container():
        st.markdown("<hr>", unsafe_allow_html=True)
        st.subheader("執行結果")
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("執行筆數", r["processed"])
        c2.metric("成功",     r["success"])
        c3.metric("失敗",     r["failed"])
        c4.metric("略過",     r["skipped"])
        c5.metric("回寫筆數", r["updated_orders"])

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

# ── 執行邏輯 ─────────────────────────────────────────────
if run:
    st.session_state.logs = []
    st.session_state.last_result = None
    st.session_state.ran = False
    result_area.empty()
    render_logs()

    if not email or not password:
        err = "請先輸入 Email / Password"
        ui_log(err)
        result = {**DEFAULT_RESULT, "failed": 1, "errors": [err]}
        st.session_state.last_result = result
        st.session_state.ran = True
        render_result(result)
    else:
        try:
            ui_log("===== 開始執行 =====")
            if mode == "By Google Sheet 列號":
                if not hasattr(memo, "main"): raise RuntimeError("memo.py 缺少 main()")
                result = memo.main(row_spec=row_spec, force=force, ui_logger=ui_log)
            elif mode == "By 電話":
                if not hasattr(memo, "main_by_phone"): raise RuntimeError("memo.py 缺少 main_by_phone()")
                result = memo.main_by_phone(phone=phone, ui_logger=ui_log)
            else:
                if not hasattr(memo, "main_by_conditions"): raise RuntimeError("memo.py 缺少 main_by_conditions()")
                result = memo.main_by_conditions(date_s=date_s, limit=int(limit), ui_logger=ui_log)

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
