# -*- coding: utf-8 -*-
import streamlit as st
import memo

st.set_page_config(page_title="Memo 自動回填系統", layout="wide")

st.title("📋 Memo 自動回填系統")

# ========================
# Session state
# ========================
if "logs" not in st.session_state:
    st.session_state.logs = []

if "last_result" not in st.session_state:
    st.session_state.last_result = {
        "processed": 0,
        "success": 0,
        "failed": 0,
        "skipped": 0,
        "updated_orders": 0,
        "errors": [],
    }


def reset_run_state():
    st.session_state.logs = []
    st.session_state.last_result = {
        "processed": 0,
        "success": 0,
        "failed": 0,
        "skipped": 0,
        "updated_orders": 0,
        "errors": [],
    }


# ========================
# Environment
# ========================
st.subheader("1. 選擇環境")

env_option = st.radio(
    "環境",
    ["prod", "dev"],
    horizontal=True,
    index=0,
)

memo.set_env(env_option)
st.caption(f"目前環境：{env_option}")

# ========================
# Login credentials
# ========================
st.subheader("2. 登入帳密")

email = st.text_input("Email", value="", key="login_email")
password = st.text_input("Password", value="", type="password", key="login_password")

memo.set_runtime_credentials(email, password)

# ========================
# Mode
# ========================
st.subheader("3. 選擇模式")

mode = st.radio(
    "處理方式",
    [
        "Google Sheet列號模式",
        "直接輸入電話號碼模式",
        "搜尋條件模式",
    ],
    horizontal=True,
)

row_spec = ""
phone = ""
date_s = ""
limit = 10
force = False

if mode == "Google Sheet列號模式":
    row_spec = st.text_input("列號", "2,3,5-8")
    force = st.checkbox("強制重跑", value=False)

elif mode == "直接輸入電話號碼模式":
    phone = st.text_input("電話號碼", "")

elif mode == "搜尋條件模式":
    c1, c2 = st.columns(2)
    with c1:
        date_s = st.text_input("訂購日期 YYYY/MM/DD", "2026/04/09")
    with c2:
        limit = st.number_input("每次處理筆數", min_value=1, max_value=100, value=5)

    st.caption("條件：付款狀態 = 已付款，服務狀態 = 未處理")

run = st.button("🚀 執行", use_container_width=True)

# ========================
# Execution logs
# ========================
st.subheader("4. 執行過程")
log_placeholder = st.empty()


def render_logs():
    if st.session_state.logs:
        log_placeholder.code("\n".join(st.session_state.logs[-3000:]))
    else:
        log_placeholder.code("尚未執行")


def ui_log(msg: str):
    st.session_state.logs.append(str(msg))
    log_placeholder.code("\n".join(st.session_state.logs[-3000:]))


render_logs()

# ========================
# Result
# ========================
st.subheader("5. 執行結果")
metrics_placeholder = st.empty()
errors_placeholder = st.empty()


def render_result():
    last = st.session_state.last_result

    with metrics_placeholder.container():
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("執行筆數", last.get("processed", 0))
        m2.metric("成功", last.get("success", 0))
        m3.metric("失敗", last.get("failed", 0))
        m4.metric("略過", last.get("skipped", 0))
        m5.metric("回寫筆數", last.get("updated_orders", 0))

    with errors_placeholder.container():
        errors = last.get("errors", [])
        if errors:
            st.subheader("6. 失敗明細")
            for err in errors:
                st.error(err)


render_result()

# ========================
# Run
# ========================
if run:
    reset_run_state()
    render_logs()
    render_result()

    ui_log("準備執行中...")

    if not email or not password:
        err = "請先輸入 Email / Password"
        ui_log(err)
        st.session_state.last_result = {
            "processed": 0,
            "success": 0,
            "failed": 1,
            "skipped": 0,
            "updated_orders": 0,
            "errors": [err],
        }
        render_result()
        st.error(err)
    else:
        try:
            if mode == "Google Sheet列號模式":
                result = memo.main(
                    row_spec=row_spec,
                    force=force,
                    ui_logger=ui_log,
                )
            elif mode == "直接輸入電話號碼模式":
                result = memo.main_by_phone(
                    phone=phone,
                    ui_logger=ui_log,
                )
            else:
                result = memo.main_by_conditions(
                    date_s=date_s,
                    limit=int(limit),
                    ui_logger=ui_log,
                )

            st.session_state.last_result = result
            render_result()
            ui_log("執行完成")
            st.success("執行完成")

        except Exception as e:
            err = f"執行失敗：{e}"
            ui_log(err)
            st.session_state.last_result = {
                "processed": 0,
                "success": 0,
                "failed": 1,
                "skipped": 0,
                "updated_orders": 0,
                "errors": [str(e)],
            }
            render_result()
            st.error(err)
