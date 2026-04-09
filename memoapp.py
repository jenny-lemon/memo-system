import streamlit as st
import memo

st.set_page_config(
    page_title="Memo 自動補備註系統",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
    /* 整體頁面風格 */
    .block-container {
        padding-top: 2rem;
        padding-bottom: 3rem;
        max-width: 960px;
    }

    /* 頁首區塊 */
    .app-header {
        background: linear-gradient(135deg, #1a1a2e 0%, #16213e 60%, #0f3460 100%);
        border-radius: 16px;
        padding: 2rem 2.5rem;
        margin-bottom: 2rem;
        color: white;
    }
    .app-header h1 {
        font-size: 1.8rem;
        font-weight: 700;
        margin: 0 0 0.3rem 0;
        letter-spacing: -0.5px;
    }
    .app-header p {
        font-size: 0.95rem;
        opacity: 0.75;
        margin: 0;
    }

    /* 環境資訊 Badge */
    .env-badge {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        background: rgba(255,255,255,0.12);
        border: 1px solid rgba(255,255,255,0.2);
        border-radius: 8px;
        padding: 6px 14px;
        font-size: 0.85rem;
        color: white;
        margin-right: 10px;
        margin-top: 1rem;
    }
    .env-badge span.label {
        opacity: 0.65;
        font-size: 0.75rem;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }

    /* 卡片 */
    .card {
        background: white;
        border: 1px solid #e8e8e8;
        border-radius: 12px;
        padding: 1.5rem;
        margin-bottom: 1.2rem;
        box-shadow: 0 1px 4px rgba(0,0,0,0.05);
    }
    .card-title {
        font-size: 0.75rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 1px;
        color: #888;
        margin-bottom: 1rem;
    }

    /* Metric cards */
    [data-testid="metric-container"] {
        background: #f7f8fa;
        border: 1px solid #e8e8e8;
        border-radius: 12px;
        padding: 1rem 1.2rem;
    }
    [data-testid="metric-container"] label {
        font-size: 0.78rem !important;
        color: #888 !important;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    [data-testid="metric-container"] [data-testid="stMetricValue"] {
        font-size: 2rem !important;
        font-weight: 700 !important;
        color: #1a1a2e !important;
    }

    /* 執行按鈕 */
    .stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #0f3460, #1a1a2e);
        border: none;
        border-radius: 10px;
        padding: 0.65rem 2rem;
        font-size: 1rem;
        font-weight: 600;
        letter-spacing: 0.3px;
        transition: opacity 0.15s ease;
        width: 100%;
    }
    .stButton > button[kind="primary"]:hover {
        opacity: 0.88;
    }

    /* Log 輸出 */
    .log-box {
        background: #0d1117;
        border-radius: 10px;
        border: 1px solid #30363d;
        padding: 1rem 1.2rem;
        font-family: 'SF Mono', 'Fira Code', monospace;
        font-size: 0.82rem;
        color: #c9d1d9;
        min-height: 180px;
        max-height: 420px;
        overflow-y: auto;
        white-space: pre-wrap;
        word-break: break-all;
        line-height: 1.6;
    }

    /* Section 標題 */
    .section-label {
        font-size: 0.72rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 1px;
        color: #aaa;
        margin-bottom: 0.5rem;
    }

    /* Divider */
    hr.divider {
        border: none;
        border-top: 1px solid #f0f0f0;
        margin: 1.5rem 0;
    }

    /* Expander 美化 */
    [data-testid="stExpander"] {
        border: 1px solid #e8e8e8 !important;
        border-radius: 10px !important;
        overflow: hidden;
    }

    /* 隱藏 Streamlit footer */
    footer { visibility: hidden; }
    #MainMenu { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ── 頁首 ───────────────────────────────────────────────
st.markdown(f"""
<div class="app-header">
    <h1>📋 Memo 自動補備註系統</h1>
    <p>批次比對上一筆已處理訂單，回填客服備註，並將未處理訂單改為已處理</p>
    <div style="margin-top: 1rem; display: flex; flex-wrap: wrap; gap: 8px;">
        <div class="env-badge">
            <span class="label">ENV</span>
            <strong>{memo.ENV_NAME}</strong>
        </div>
        <div class="env-badge">
            <span class="label">BASE URL</span>
            <strong>{memo.BASE_URL}</strong>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# ── 主體三欄佈局 ────────────────────────────────────────
left_col, right_col = st.columns([1, 1.6], gap="large")

with left_col:
    # 處理設定
    st.markdown('<div class="section-label">處理設定</div>', unsafe_allow_html=True)
    row_spec = st.text_input(
        "要處理的列",
        value="2,3,5-8",
        placeholder="例如：2,3,5-8",
        help="可輸入單列、多列、區間，例如：2,3,5-8",
    )
    force = st.checkbox("強制重跑（忽略 V 欄已有值）", value=False)

    st.markdown("<hr class='divider'>", unsafe_allow_html=True)

    # 帳號設定
    st.markdown('<div class="section-label">帳號設定</div>', unsafe_allow_html=True)

    with st.expander("🏙 台北帳號", expanded=True):
        taipei_email = st.text_input(
            "Email",
            value=memo.accounts.ACCOUNTS.get("台北", {}).get("email", ""),
            key="taipei_email",
            placeholder="taipei@example.com",
        )
        taipei_password = st.text_input(
            "Password",
            value=memo.accounts.ACCOUNTS.get("台北", {}).get("password", ""),
            type="password",
            key="taipei_password",
            placeholder="••••••••",
        )

    with st.expander("🏙 台中帳號", expanded=False):
        taichung_email = st.text_input(
            "Email",
            value=memo.accounts.ACCOUNTS.get("台中", {}).get("email", ""),
            key="taichung_email",
            placeholder="taichung@example.com",
        )
        taichung_password = st.text_input(
            "Password",
            value=memo.accounts.ACCOUNTS.get("台中", {}).get("password", ""),
            type="password",
            key="taichung_password",
            placeholder="••••••••",
        )

    st.markdown("<br>", unsafe_allow_html=True)
    run_btn = st.button("🚀 開始執行", type="primary", use_container_width=True)

with right_col:
    # 執行狀態
    st.markdown('<div class="section-label">執行狀態</div>', unsafe_allow_html=True)

    metric_col1, metric_col2, metric_col3, metric_col4 = st.columns(4)
    m1 = metric_col1.empty()
    m2 = metric_col2.empty()
    m3 = metric_col3.empty()
    m4 = metric_col4.empty()

    def render_metrics(processed="-", success="-", failed="-", skipped="-"):
        m1.metric("執行", processed)
        m2.metric("完成", success)
        m3.metric("失敗", failed)
        m4.metric("略過", skipped)

    render_metrics()

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<div class="section-label">執行記錄</div>', unsafe_allow_html=True)
    log_placeholder = st.empty()
    log_placeholder.markdown('<div class="log-box">等待執行...</div>', unsafe_allow_html=True)

    error_placeholder = st.empty()

# ── session state ────────────────────────────────────────
if "live_logs" not in st.session_state:
    st.session_state.live_logs = []

def ui_logger(message: str):
    st.session_state.live_logs.append(message)
    log_content = "\n".join(st.session_state.live_logs[-300:])
    log_placeholder.markdown(
        f'<div class="log-box">{log_content}</div>',
        unsafe_allow_html=True,
    )

# ── 執行邏輯 ─────────────────────────────────────────────
if run_btn:
    st.session_state.live_logs = []
    log_placeholder.markdown(
        '<div class="log-box">準備執行中...</div>',
        unsafe_allow_html=True,
    )
    error_placeholder.empty()

    if "台北" in memo.accounts.ACCOUNTS:
        memo.accounts.ACCOUNTS["台北"]["email"] = taipei_email
        memo.accounts.ACCOUNTS["台北"]["password"] = taipei_password
    if "台中" in memo.accounts.ACCOUNTS:
        memo.accounts.ACCOUNTS["台中"]["email"] = taichung_email
        memo.accounts.ACCOUNTS["台中"]["password"] = taichung_password

    try:
        result = memo.main(
            row_spec=row_spec,
            force=force,
            ui_logger=ui_logger,
        )
        render_metrics(
            processed=result.get("processed", 0),
            success=result.get("success", 0),
            failed=result.get("failed", 0),
            skipped=result.get("skipped", 0),
        )

        if result.get("errors"):
            with error_placeholder.container():
                st.markdown('<div class="section-label" style="margin-top:1.5rem;">失敗明細</div>', unsafe_allow_html=True)
                for err in result["errors"]:
                    st.error(err)

        st.toast("✅ 執行完成！", icon="✅")

    except Exception as e:
        st.toast(f"執行失敗：{e}", icon="❌")
        log_placeholder.markdown(
            f'<div class="log-box" style="color:#f85149;">❌ 執行失敗：{e}</div>',
            unsafe_allow_html=True,
        )
