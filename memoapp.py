# -*- coding: utf-8 -*-
import streamlit as st
import memo

st.set_page_config(page_title="Memo 自動回填系統", layout="wide")

st.title("📋 Memo 自動回填系統")

# ========================
# Session state
# ========================
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
    st.session_state.last_result = DEFAULT_RESULT.copy()


def reset_run_state():
    st.session_state.logs = []
    st.session_state.last_result = DEFAULT_RESULT.copy()


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

if hasattr(memo, "set_env"):
    memo.set_env(env_option)

base_url = getattr(memo, "BASE_URL", "")
if base_url:
    st.caption(f"目前環境：{env_option} | {base_url}")
else:
    st.caption(f"目前環境：{env_option}")


# ========================
# Credentials
# ========================
st.subheader("2. 登入帳密")

email = st.text_input("Email", value="", key="login_email")
password = st.text_input("Password", value="", type="password", key="login_password")

if hasattr(memo, "set_runtime_credentials"):
    memo.set_runtime_credentials(email, password)
else:
    setattr(memo, "RUNTIME_EMAIL", email)
    setattr(memo, "RUNTIME_PASSWORD", password)


# ========================
# Mode
# ========================
st.subheader("3. 選擇模式")

mode = st.radio(
    "處理方式",
    [
        "By Google Sheet 列號",
        "By 電話",
        "By 搜尋條件",
    ],
    horizontal=True,
)

row_spec = ""
phone = ""
date_s = ""
limit = 5
force = False

if mode == "By Google Sheet 列號":
    row_spec = st.text_input("列號", "2,3,5-8")
    force = st.checkbox("強制重跑", value=False)

elif mode == "By 電話":
    phone = st.text_input("電話號碼", "")

elif mode == "By 搜尋條件":
    c1, c2 = st.columns(2)
    with c1:
        date_s = st.text_input("訂購日期 YYYY/MM/DD", "")
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
# Result UI
# ========================
metrics_placeholder = st.empty()
errors_placeholder = st.empty()


def normalize_result(result):
    base = DEFAULT_RESULT.copy()
    if isinstance(result, dict):
        base.update(result)
    if "updated_orders" not in base:
        base["updated_orders"] = 0
    if "errors" not in base or base["errors"] is None:
        base["errors"] = []
    return base


def render_result():
    last = normalize_result(st.session_state.last_result)

    st.subheader("5. 執行結果")
    with metrics_placeholder.container():
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("執行筆數", last.get("processed", 0))
        c2.metric("成功", last.get("success", 0))
        c3.metric("失敗", last.get("failed", 0))
        c4.metric("略過", last.get("skipped", 0))
        c5.metric("回寫筆數", last.get("updated_orders", 0))

    with errors_placeholder.container():
        errors = last.get("errors", [])
        if errors:
            st.subheader("6. 錯誤明細")
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
            ui_log("===== 開始執行 =====")

            if mode == "By Google Sheet 列號":
                if not hasattr(memo, "main"):
                    raise RuntimeError("memo.py 缺少 main()")
                result = memo.main(
                    row_spec=row_spec,
                    force=force,
                    ui_logger=ui_log,
                )

            elif mode == "By 電話":
                if not hasattr(memo, "main_by_phone"):
                    raise RuntimeError("memo.py 缺少 main_by_phone()")
                result = memo.main_by_phone(
                    phone=phone,
                    ui_logger=ui_log,
                )

            else:
                if not hasattr(memo, "main_by_conditions"):
                    raise RuntimeError("memo.py 缺少 main_by_conditions()")
                result = memo.main_by_conditions(
                    date_s=date_s,
                    limit=int(limit),
                    ui_logger=ui_log,
                )

            st.session_state.last_result = normalize_result(result)
            ui_log("===== 執行完成 =====")
            render_result()
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
