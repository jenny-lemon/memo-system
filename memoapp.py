import streamlit as st
import memo

st.set_page_config(page_title="Memo System", layout="wide")

st.title("📋 Memo 自動補備註系統")
st.caption("比對上一筆已處理訂單的客服備註，回填到同電話同地址的未處理訂單，並改成已處理")

# -------------------------
# 環境選擇
# -------------------------
st.subheader("環境資訊")

env_options = {
    "正式機 prod": "prod",
    "測試機 dev": "dev",
}

default_env_label = "測試機 dev" if memo.ENV_NAME == "dev" else "正式機 prod"

selected_env_label = st.radio(
    "請選擇環境",
    options=list(env_options.keys()),
    index=list(env_options.keys()).index(default_env_label),
    horizontal=True,
)

selected_env = env_options[selected_env_label]

# 依畫面選擇即時覆寫 memo 的環境
memo.ENV_NAME = selected_env
if selected_env == "dev":
    memo.BASE_URL = memo.BASE_URL_DEV.rstrip("/")
else:
    memo.BASE_URL = memo.BASE_URL_PROD.rstrip("/")

memo.LOGIN_URL = f"{memo.BASE_URL}/login"
memo.PURCHASE_URL = f"{memo.BASE_URL}/purchase"

c1, c2 = st.columns(2)
c1.info(f"ENV：{memo.ENV_NAME}")
c2.info(f"BASE_URL：{memo.BASE_URL}")

# -------------------------
# 處理設定
# -------------------------
st.subheader("處理設定")
row_spec = st.text_input("要處理的列", value="2,3,5-8", help="例如：2,3,5-8")
force = st.checkbox("強制重跑（忽略 V 欄已有值）", value=False)

# -------------------------
# 登入帳密
# -------------------------
st.subheader("登入帳密")

with st.expander("台北帳號", expanded=True):
    taipei_email = st.text_input(
        "台北 Email",
        value=memo.accounts.ACCOUNTS.get("台北", {}).get("email", ""),
        key="taipei_email",
    )
    taipei_password = st.text_input(
        "台北 Password",
        value=memo.accounts.ACCOUNTS.get("台北", {}).get("password", ""),
        type="password",
        key="taipei_password",
    )

with st.expander("台中帳號", expanded=False):
    taichung_email = st.text_input(
        "台中 Email",
        value=memo.accounts.ACCOUNTS.get("台中", {}).get("email", ""),
        key="taichung_email",
    )
    taichung_password = st.text_input(
        "台中 Password",
        value=memo.accounts.ACCOUNTS.get("台中", {}).get("password", ""),
        type="password",
        key="taichung_password",
    )

# -------------------------
# 執行狀態
# -------------------------
st.subheader("執行狀態")
log_box = st.empty()

m1, m2, m3, m4 = st.columns(4)

if "live_logs" not in st.session_state:
    st.session_state.live_logs = []


def ui_logger(message: str):
    st.session_state.live_logs.append(message)
    log_box.code("\n".join(st.session_state.live_logs[-400:]))


if st.button("🚀 開始執行", type="primary"):
    st.session_state.live_logs = []
    log_box.code("準備執行中...")

    # 覆寫帳密到 memo.accounts
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

        m1.metric("執行筆數", result.get("processed", 0))
        m2.metric("完成筆數", result.get("success", 0))
        m3.metric("失敗筆數", result.get("failed", 0))
        m4.metric("略過筆數", result.get("skipped", 0))

        if result.get("errors"):
            st.subheader("失敗明細")
            for err in result["errors"]:
                st.error(err)

        st.success("執行完成")

    except Exception as e:
        st.error(f"執行失敗：{e}")
