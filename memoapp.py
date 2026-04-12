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
    max-width: 1180px !important;
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

/* 按鈕 */
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
if "is_running" not in st.session_state:
    st.session_state.is_running = False
if "is_logged_in" not in st.session_state:
    st.session_state.is_logged_in = False
if "preview_rows" not in st.session_state:
    st.session_state.preview_rows = []
if "last_mode" not in st.session_state:
    st.session_state.last_mode = ""
if "login_identity" not in st.session_state:
    st.session_state.login_identity = ""


def normalize_result(r):
    base = DEFAULT_RESULT.copy()
    if isinstance(r, dict):
        base.update(r)
    if not isinstance(base.get("errors"), list):
        base["errors"] = []
    return base


def sec(title):
    st.markdown(f'<p class="sec-label">{title}</p>', unsafe_allow_html=True)


def render_result(result):
    r = normalize_result(result)
    with result_container:
        st.markdown("<hr>", unsafe_allow_html=True)
        sec("執行結果")
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("執行筆數", r["processed"])
        c2.metric("成功", r["success"])
        c3.metric("失敗", r["failed"])
        c4.metric("略過", r["skipped"])
        c5.metric("回寫筆數", r["updated_orders"])

        if r["errors"]:
            with st.expander(f"⚠️ 錯誤明細（{len(r['errors'])} 筆）", expanded=True):
                for i, err in enumerate(r["errors"], 1):
                    st.markdown(f"**{i}.** {err}")
        elif r["processed"] > 0:
            st.success(f"✅ 全部完成，共處理 **{r['processed']}** 筆，成功 **{r['success']}** 筆。")
        else:
            st.info("執行完成，無資料被處理。")


def ui_log(msg):
    st.session_state.logs.append(str(msg))
    try:
        log_box.code("\n".join(st.session_state.logs[-3000:]))
    except Exception:
        pass


def reset_preview_if_mode_changed(current_mode):
    if st.session_state.last_mode != current_mode:
        st.session_state.preview_rows = []
        st.session_state.last_mode = current_mode


def safe_get(row, *keys, default=""):
    for k in keys:
        if k in row and row.get(k) is not None:
            return row.get(k)
    return default


def clear_result():
    st.session_state.result = None


def render_group_rows(block_rows, key_prefix, selected_ids):
    header = st.columns([0.8, 1.4, 1.2, 2.7, 1.2, 1.0, 1.0, 1.0])
    header[0].markdown("**選取**")
    header[1].markdown("**訂單編號**")
    header[2].markdown("**姓名**")
    header[3].markdown("**地址**")
    header[4].markdown("**日期**")
    header[5].markdown("**同地址**")
    header[6].markdown("**不同地址**")
    header[7].markdown("**同址紀錄**")

    for row in block_rows:
        order_id = str(safe_get(row, "order_id", "id", "orderNo", default="")).strip()
        customer_name = safe_get(row, "customer_name", "name", "customer", default="")
        address = safe_get(row, "address", "service_address", default="")
        service_date = safe_get(row, "service_date", "date", "purchase_date", default="")
        same_count = int(safe_get(row, "same_address_count", default=0) or 0)
        diff_count = int(safe_get(row, "different_address_count", default=0) or 0)
        has_same = bool(safe_get(row, "has_same_address_history", default=False))

        cols = st.columns([0.8, 1.4, 1.2, 2.7, 1.2, 1.0, 1.0, 1.0])

        checked = cols[0].checkbox(
            "選取",
            key=f"{key_prefix}_{order_id}",
            label_visibility="collapsed",
            value=st.session_state.get(f"{key_prefix}_{order_id}", st.session_state.get(f"pick_{order_id}", False)),
        )

        st.session_state[f"pick_{order_id}"] = checked

        cols[1].write(order_id)
        cols[2].write(customer_name)
        cols[3].write(address)
        cols[4].write(service_date)
        cols[5].write(f"{same_count} 筆")
        cols[6].write(f"{diff_count} 筆")
        cols[7].write("有" if has_same else "無")

        if checked and order_id:
            selected_ids.append(order_id)


def render_preview_list(rows):
    sec("查詢結果")

    if not rows:
        st.info("查無資料")
        return []

    rows = sorted(
        rows,
        key=lambda x: (
            0 if x.get("has_same_address_history") else 1,
            -(x.get("same_address_count", 0) or 0),
            str(x.get("service_date", "")),
            str(x.get("order_id", "")),
        )
    )

    same_rows = [r for r in rows if r.get("has_same_address_history")]
    diff_rows = [r for r in rows if not r.get("has_same_address_history")]

    ctl1, ctl2, ctl3 = st.columns([1, 1, 4])
    with ctl1:
        if st.button("全選", use_container_width=True, key="select_all_rows"):
            for row in rows:
                oid = str(safe_get(row, "order_id", "id", "orderNo", default="")).strip()
                if oid:
                    st.session_state[f"pick_{oid}"] = True
                    st.session_state[f"same_{oid}"] = True
                    st.session_state[f"diff_{oid}"] = True
            st.rerun()

    with ctl2:
        if st.button("全不選", use_container_width=True, key="unselect_all_rows"):
            for row in rows:
                oid = str(safe_get(row, "order_id", "id", "orderNo", default="")).strip()
                if oid:
                    st.session_state[f"pick_{oid}"] = False
                    st.session_state[f"same_{oid}"] = False
                    st.session_state[f"diff_{oid}"] = False
            st.rerun()

    st.caption(f"有同地址 {len(same_rows)} 筆｜無同地址 {len(diff_rows)} 筆")

    selected_ids = []

    if same_rows:
        st.markdown("#### 有同地址紀錄")
        render_group_rows(same_rows, "same", selected_ids)
    else:
        st.markdown("#### 有同地址紀錄")
        st.caption("沒有資料")

    st.markdown("#### 無同地址紀錄")
    if diff_rows:
        render_group_rows(diff_rows, "diff", selected_ids)
    else:
        st.caption("沒有資料")

    st.caption(f"目前勾選 {len(selected_ids)} 筆")
    return selected_ids


st.title("📋 Memo 自動回填系統")

# ── 登入區 ───────────────────────────────────────────────
sec("登入")
col_e, col_p, col_env, col_login = st.columns([3.2, 3.2, 1.2, 1.2])
with col_e:
    email = st.text_input("Email")
with col_p:
    password = st.text_input("Password", type="password")
with col_env:
    env_option = st.selectbox("環境", ["prod", "dev"], index=0)
with col_login:
    st.markdown("<div style='height:26px'></div>", unsafe_allow_html=True)
    login_clicked = st.button(
        "Login",
        use_container_width=True,
        disabled=st.session_state.is_running,
    )

memo.set_env(env_option)

if login_clicked:
    st.session_state.logs = []
    clear_result()
    try:
        if not email or not password:
            st.error("請先輸入 Email / Password")
            st.stop()

        st.session_state.is_running = True
        ui_log("===== 開始登入 =====")

        memo.set_env(env_option)
        memo.set_runtime_credentials(email, password)

        with st.spinner("登入中，請稍候…"):
            memo.login(ui_logger=ui_log)

        st.session_state.is_logged_in = True
        st.session_state.login_identity = email
        ui_log("✅ Login 成功")
        st.success("登入成功")

    except Exception as e:
        st.session_state.is_logged_in = False
        st.session_state.login_identity = ""
        ui_log(f"❌ Login 失敗：{e}")
        st.error(f"登入失敗：{e}")

    finally:
        st.session_state.is_running = False

if st.session_state.is_logged_in:
    st.caption(f"目前已登入：{st.session_state.login_identity}")
else:
    st.info("請先登入後再查詢或執行。")

st.markdown("<hr>", unsafe_allow_html=True)

# ── 處理模式 ──────────────────────────────────────────────
sec("處理模式")
mode = st.radio(
    "",
    ["By Google Sheet", "By 電話", "By 搜尋條件"],
    horizontal=True,
    label_visibility="collapsed",
)

reset_preview_if_mode_changed(mode)

row_spec = ""
phone = ""
date_mode = "服務日期"
purchase_status_name = "已付款"
limit = 5
start_date = None
end_date = None
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

st.markdown("<hr>", unsafe_allow_html=True)

# ── 查詢 / 執行按鈕 ───────────────────────────────────────
if mode == "By Google Sheet":
    c_run = st.columns(1)[0]
    with c_run:
        execute_btn = st.button(
            "🚀 執行",
            use_container_width=True,
            disabled=(st.session_state.is_running or not st.session_state.is_logged_in),
        )
    search_btn = False
else:
    c_search, c_run = st.columns(2)
    with c_search:
        search_btn = st.button(
            "🔍 查詢列表",
            use_container_width=True,
            disabled=(st.session_state.is_running or not st.session_state.is_logged_in),
        )
    with c_run:
        execute_btn = st.button(
            "🚀 執行勾選項目",
            use_container_width=True,
            disabled=(st.session_state.is_running or not st.session_state.is_logged_in),
        )

# ── 執行過程 ──────────────────────────────────────────────
with st.expander("📄 執行過程", expanded=True):
    log_box = st.empty()
    log_box.code("\n".join(st.session_state.logs[-3000:]) if st.session_state.logs else "尚未執行")

# ── 查詢結果 ──────────────────────────────────────────────
selected_ids = []
if mode in ["By 電話", "By 搜尋條件"] and st.session_state.preview_rows:
    st.markdown("<hr>", unsafe_allow_html=True)
    selected_ids = render_preview_list(st.session_state.preview_rows)

# ── 執行結果 ──────────────────────────────────────────────
result_container = st.container()
if st.session_state.result is not None:
    render_result(st.session_state.result)

# ── 查詢邏輯 ──────────────────────────────────────────────
if search_btn:
    clear_result()

    if not st.session_state.is_logged_in:
        st.warning("請先登入")
        st.stop()

    try:
        st.session_state.is_running = True
        st.session_state.logs = []
        st.session_state.preview_rows = []
        log_box.code("尚未執行")

        ui_log("===== 開始查詢 =====")

        with st.spinner("查詢中，請稍候…"):
            if mode == "By 電話":
                if not phone.strip():
                    raise ValueError("請先輸入電話號碼")

                preview_rows = memo.preview_by_phone(
                    phone=phone.strip(),
                    ui_logger=ui_log,
                )

            elif mode == "By 搜尋條件":
                start_text = start_date.strftime("%Y/%m/%d") if start_date else ""
                end_text = end_date.strftime("%Y/%m/%d") if end_date else ""

                preview_rows = memo.preview_by_conditions(
                    date_mode=date_mode,
                    date_start=start_text,
                    date_end=end_text,
                    purchase_status_name=purchase_status_name,
                    limit=int(limit),
                    ui_logger=ui_log,
                )
            else:
                preview_rows = []

        st.session_state.preview_rows = preview_rows or []
        ui_log(f"✅ 查詢完成，共 {len(st.session_state.preview_rows)} 筆")
        st.rerun()

    except Exception as e:
        ui_log(f"❌ 查詢錯誤：{e}")
        st.error(str(e))

    finally:
        st.session_state.is_running = False

# ── 執行邏輯 ──────────────────────────────────────────────
if execute_btn:
    clear_result()

    if not st.session_state.is_logged_in:
        st.warning("請先登入")
        st.stop()

    try:
        st.session_state.is_running = True

        if mode == "By Google Sheet":
            st.session_state.logs = []
            log_box.code("尚未執行")
            ui_log("===== 開始執行 =====")

            with st.spinner("執行中，請稍候…"):
                result = memo.main(
                    row_spec=row_spec,
                    force=force,
                    ui_logger=ui_log,
                )

        else:
            if not st.session_state.preview_rows:
                st.warning("請先查詢列表")
                st.stop()

            current_selected_ids = []
            for row in st.session_state.preview_rows:
                oid = str(safe_get(row, "order_id", "id", "orderNo", default="")).strip()
                if oid and st.session_state.get(f"pick_{oid}", False):
                    current_selected_ids.append(oid)

            if not current_selected_ids:
                st.warning("請先勾選要執行的資料")
                st.stop()

            st.session_state.logs = []
            log_box.code("尚未執行")
            ui_log("===== 開始執行勾選項目 =====")
            ui_log(f"勾選筆數：{len(current_selected_ids)}")

            with st.spinner("執行中，請稍候…"):
                result = memo.main_by_selected_order_ids(
                    order_ids=current_selected_ids,
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
