# -*- coding: utf-8 -*-
import streamlit as st
import memo

st.set_page_config(page_title="Memo 自動回填系統", layout="wide")

# ── 樣式 ──────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@300;400;500;600&family=DM+Sans:opsz,wght@9..40,300;9..40,400;9..40,500;9..40,600&display=swap');

html, body, [class*="css"] {
    font-family: 'DM Sans', 'Noto Sans TC', sans-serif;
}

/* 頁面標題 */
h1 { font-size: 22px !important; font-weight: 600 !important; margin-bottom: 0 !important; }

/* Section 標題統一 */
h3 { font-size: 13px !important; font-weight: 500 !important;
     color: #8892b0 !important; letter-spacing: 0.08em !important;
     text-transform: uppercase !important; margin-bottom: 8px !important; margin-top: 20px !important; }

/* 分隔線 */
hr { border-color: rgba(255,255,255,0.07) !important; margin: 18px 0 !important; }

/* Radio 按鈕組 */
[data-testid="stRadio"] > div { gap: 8px !important; }
[data-testid="stRadio"] label {
    background: rgba(255,255,255,0.04) !important;
    border: 1px solid rgba(255,255,255,0.09) !important;
    border-radius: 7px !important; padding: 6px 16px !important;
    font-size: 13.5px !important; transition: all 0.15s !important;
}
[data-testid="stRadio"] label:hover { border-color: #8ba4ff !important; }
[data-testid="stRadio"] label[data-checked="true"] {
    background: rgba(139,164,255,0.12) !important;
    border-color: #8ba4ff !important; color: #8ba4ff !important;
}

/* 執行按鈕 */
[data-testid="stButton"] > button[kind="primary"],
[data-testid="stButton"] > button {
    background: linear-gradient(135deg, #6c8fff, #a78bfa) !important;
    border: none !important; color: #fff !important; font-weight: 500 !important;
    font-size: 14px !important; border-radius: 9px !important;
    padding: 10px 0 !important; letter-spacing: 0.02em !important;
    box-shadow: 0 4px 14px rgba(108,143,255,0.3) !important;
    transition: opacity 0.2s !important;
}
[data-testid="stButton"] > button:hover { opacity: 0.88 !important; }

/* Log 區塊 */
[data-testid="stCode"] {
    background: #141622 !important; border-radius: 9px !important;
    border: 1px solid rgba(255,255,255,0.07) !important;
    font-size: 12px !important; max-height: 320px; overflow-y: auto;
}

/* Metric 卡片 */
[data-testid="stMetric"] {
    background: rgba(255,255,255,0.03) !important;
    border: 1px solid rgba(255,255,255,0.08) !important;
    border-radius: 10px !important; padding: 14px 16px !important;
}
[data-testid="stMetricLabel"] { font-size: 12px !important; color: #6b7290 !important; }
[data-testid="stMetricValue"] { font-size: 26px !important; font-weight: 600 !important; }

/* Caption */
[data-testid="stCaptionContainer"] { color: #4e5470 !important; font-size: 12px !important; }
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
mode = st.radio(
    "mode", ["By Google Sheet 列號", "By 電話", "By 搜尋條件"],
    horizontal=True, label_visibility="collapsed"
)

row_spec = phone = date_s = ""
limit = 5; force = False

if mode == "By Google Sheet 列號":
    col_r, col_f = st.columns([3, 1])
    with col_r: row_spec = st.text_input("列號（例：2,3,5-8）", "2,3,5-8")
    with col_f:
        st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
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

        # 各狀態數字
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("執行筆數", r["processed"])
        c2.metric("成功",     r["success"])
        c3.metric("失敗",     r["failed"])
        c4.metric("略過",     r["skipped"])
        c5.metric("回寫筆數", r["updated_orders"])

        # 結果說明
        if r["failed"] == 0 and r["processed"] > 0:
            st.success(f"✅ 全部執行完成，共處理 {r['processed']} 筆，成功 {r['success']} 筆。")
        elif r["failed"] > 0:
            st.warning(f"⚠️ 執行完成，但有 {r['failed']} 筆失敗，請查看下方錯誤明細。")
        else:
            st.info("執行完成，無資料被處理。")

        # 錯誤明細
        if r["errors"]:
            st.subheader("錯誤明細")
            for err in r["errors"]:
                st.error(err)

# 若上次有結果，重新渲染（頁面 rerun 後保留）
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
                if not hasattr(memo, "main"):
                    raise RuntimeError("memo.py 缺少 main()")
                result = memo.main(row_spec=row_spec, force=force, ui_logger=ui_log)

            elif mode == "By 電話":
                if not hasattr(memo, "main_by_phone"):
                    raise RuntimeError("memo.py 缺少 main_by_phone()")
                result = memo.main_by_phone(phone=phone, ui_logger=ui_log)

            else:
                if not hasattr(memo, "main_by_conditions"):
                    raise RuntimeError("memo.py 缺少 main_by_conditions()")
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
