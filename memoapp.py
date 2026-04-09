import streamlit as st
import memo
import accounts

st.set_page_config(page_title="Memo System", layout="wide")

st.title("📋 Memo 自動補備註系統")
st.caption("搜尋手機號碼 + 已付款，依相同地址取得前次客服備註，回填並將服務狀態改為已處理")

# -------------------------
# 環境選擇
# -------------------------
env_options = {
    "正式機 prod": "prod",
    "測試機 dev": "dev",
}

default_env_label = "測試機 dev" if memo.ENV_NAME == "dev" else "正式機 prod"

with st.container(border=True):
    st.subheader("1. 選擇環境")

    selected_env_label = st.radio(
        "請選擇環境",
        options=list(env_options.keys()),
        index=list(env_options.keys()).index(default_env_label),
        horizontal=True,
    )

    selected_env = env_options[selected_env_label]
    memo.set_env(selected_env)

    c1, c2 = st.columns(2)
    c1.info(f"ENV：{memo.ENV_NAME}")
    c2.info(f"BASE_URL：{memo.BASE_URL}")

# -------------------------
# 帳密
# -------------------------
with st.container(border=True):
    st.subheader("2. 輸入登入帳密")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("#### 台北")
        taipei_email = st.text_input(
            "台北 email",
            value=accounts.ACCOUNTS.get("台北", {}).get("email", ""),
            key="taipei_email",
        )
        taipei_password = st.text_input(
            "台北 password",
            value=accounts.ACCOUNTS.get("台北", {}).get("password", ""),
            type="password",
            key="taipei_password",
        )

    with col2:
        st.markdown("#### 台中")
        taichung_email = st.text_input(
            "台中 email",
            value=accounts.ACCOUNTS.get("台中", {}).get("email", ""),
            key="taichung_email",
        )
        taichung_password = st.text_input(
            "台中 password",
            value=accounts.ACCOUNTS.get("台中", {}).get("password", ""),
            type="password",
            key="taichung_password",
        )

# -------------------------
# 處理設定
# -------------------------
with st.container(border=True):
    st.subheader("3. 處理設定")

    row_spec = st.text_input(
        "處理列",
        value="2,3,5-8",
        help="例如：2,3,5-8",
    )

    force = st.checkbox("強制重跑", value=False)

    run_btn = st.button("🚀 執行", type="primary", use_container_width=False)

# -------------------------
# 執行結果區塊（放在按鈕下方）
# -------------------------
st.subheader("4. 執行結果")

metric_col1, metric_col2, metric_col3, metric_col4 = st.columns(4)
log_box = st.empty()

if "live_logs" not in st.session_state:
    st.session_state.live_logs = []

def ui_logger(message: str):
    st.session_state.live_logs.append(message)
    log_box.code("\n".join(st.session_state.live_logs[-500:]))

if run_btn:
    st.session_state.live_logs = []
    log_box.code("準備執行中...")

    if "台北" in accounts.ACCOUNTS:
        accounts.ACCOUNTS["台北"]["email"] = taipei_email
        accounts.ACCOUNTS["台北"]["password"] = taipei_password

    if "台中" in accounts.ACCOUNTS:
        accounts.ACCOUNTS["台中"]["email"] = taichung_email
        accounts.ACCOUNTS["台中"]["password"] = taichung_password

    try:
        result = memo.main(
            row_spec=row_spec,
            force=force,
            ui_logger=ui_logger,
        )

        metric_col1.metric("執行筆數", result.get("processed", 0))
        metric_col2.metric("完成筆數", result.get("success", 0))
        metric_col3.metric("失敗筆數", result.get("failed", 0))
        metric_col4.metric("略過筆數", result.get("skipped", 0))

        if result.get("errors"):
            st.markdown("#### 失敗明細")
            for err in result["errors"]:
                st.error(err)

        st.success("執行完成")

    except Exception as e:
        st.error(f"執行失敗：{e}")
