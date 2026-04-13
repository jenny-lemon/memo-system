# memoapp.py
# -*- coding: utf-8 -*-
import streamlit as st
import memo

st.set_page_config(page_title="Memo 自動回填系統", layout="wide")

st.markdown("""
<style>
html, body, [class*="css"] {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Noto Sans TC", sans-serif;
}
.block-container {
    padding-top: 1.2rem !important;
    padding-bottom: 2rem !important;
    max-width: 1320px !important;
}
h1 {
    font-size: 20px !important;
    font-weight: 700 !important;
    margin-bottom: 14px !important;
}
.sec-label {
    font-size: 12px;
    font-weight: 700;
    color: #7a7a7a;
    letter-spacing: .04em;
    margin-bottom: 8px;
}
[data-testid="stButton"] > button {
    background: #111 !important;
    color: #fff !important;
    border: none !important;
    border-radius: 10px !important;
    font-size: 15px !important;
    font-weight: 700 !important;
    padding: 10px 0 !important;
}
[data-testid="stButton"] > button:hover {
    background: #333 !important;
}
[data-testid="stCode"] {
    font-size: 12px !important;
    border-radius: 10px !important;
}
[data-testid="stMetric"] {
    background: #fff;
    border: 1px solid #ececec;
    border-radius: 12px;
    padding: 10px;
}
.small-muted {
    color: #888;
    font-size: 12px;
}
.preview-card {
    border: 1px solid #ececec;
    border-radius: 14px;
    padding: 12px 14px;
    margin-bottom: 10px;
    background: #fff;
}
.preview-ok {
    border-left: 5px solid #16a34a;
}
.preview-ng {
    border-left: 5px solid #d4d4d8;
}
.preview-title {
    font-size: 16px;
    font-weight: 700;
    margin-bottom: 6px;
}
.preview-sub {
    color: #555;
    font-size: 13px;
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

DEFAULT_STATE = {
    "logs": [],
    "result": None,
    "is_running": False,
    "is_logged_in": False,
    "preview_rows": [],
    "last_mode": "",
    "login_identity": "",
    "sheet_summary": None,
}

for k, v in DEFAULT_STATE.items():
    if k not in st.session_state:
        st.session_state[k] = v

st.session_state.is_running = False


def sec(title):
    st.markdown(f'<p class="sec-label">{title}</p>', unsafe_allow_html=True)


def normalize_result(r):
    base = DEFAULT_RESULT.copy()
    if isinstance(r, dict):
        base.update(r)
    if not isinstance(base.get("errors"), list):
        base["errors"] = []
    return base


def render_result(result):
    r = normalize_result(result)
    with result_container:
        st.markdown("---")
        sec("5. 執行結果")
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
            st.success(f"✅ 全部完成，共處理 {r['processed']} 筆，成功 {r['success']} 筆。")
        else:
            st.info("執行完成，無資料被處理。")


def ui_log(msg):
    st.session_state.logs.append(str(msg))
    try:
        log_box.code("\n".join(st.session_state.logs[-3000:]))
    except Exception:
        pass


def safe_get(row, *keys, default=""):
    for k in keys:
        if k in row and row.get(k) is not None:
            return row.get(k)
    return default


def clear_pick_states():
    keys_to_delete = []
    for k in list(st.session_state.keys()):
        if k.startswith("pick_"):
            keys_to_delete.append(k)
    for k in keys_to_delete:
        del st.session_state[k]


def reset_before_action(clear_preview=True, clear_selection=True):
    st.session_state.logs = []
    st.session_state.result = None
    if clear_preview:
        st.session_state.preview_rows = []
        st.session_state.sheet_summary = None
    if clear_selection:
        clear_pick_states()
    try:
        log_box.code("尚未執行")
    except Exception:
        pass


def reset_before_execute_keep_preview():
    st.session_state.logs = []
    st.session_state.result = None
    try:
        log_box.code("尚未執行")
    except Exception:
        pass


def reset_mode_state_if_changed(current_mode):
    if st.session_state.last_mode != current_mode:
        st.session_state.preview_rows = []
        st.session_state.sheet_summary = None
        clear_pick_states()
        st.session_state.last_mode = current_mode


def render_preview_blocks(rows):
    sec("3. 查詢結果預覽")

    if not rows:
        st.info("查無資料")
        return []

    can_rows = [r for r in rows if r.get("can_autofill")]
    no_rows = [r for r in rows if not r.get("can_autofill")]

    m1, m2, m3 = st.columns(3)
    m1.metric("查詢總筆數", len(rows))
    m2.metric("可自動回填", len(can_rows))
    m3.metric("無可參照來源", len(no_rows))

    st.info("每一列是『目標訂單』；若有來源訂單，代表已找到最近一筆同地址＋已付款＋已處理＋有備註的來源。")

    selected_ids = []

    def render_section(title, items, section_key, default_checked):
        st.markdown(f"### {title}")
        if not items:
            st.caption("沒有資料")
            return

        c1, c2, c3 = st.columns([1, 1, 4])
        with c1:
            if st.button("本區全選", key=f"sel_{section_key}", use_container_width=True):
                for row in items:
                    oid = str(row.get("order_id", "")).strip()
                    if oid:
                        st.session_state[f"pick_{oid}"] = True
        with c2:
            if st.button("本區全不選", key=f"unsel_{section_key}", use_container_width=True):
                for row in items:
                    oid = str(row.get("order_id", "")).strip()
                    if oid:
                        st.session_state[f"pick_{oid}"] = False
        with c3:
            st.caption(f"本區共 {len(items)} 筆")

        for row in items:
            order_id = str(row.get("order_id", "")).strip()
            checked = st.checkbox(
                f"選取 {order_id}",
                key=f"pick_{order_id}",
                value=st.session_state.get(f"pick_{order_id}", default_checked),
                label_visibility="collapsed"
            )

            card_cls = "preview-card preview-ok" if row.get("can_autofill") else "preview-card preview-ng"

            target_name = row.get("customer_name", "")
            phone = row.get("phone", "")
            address = row.get("address", "")
            service_date = row.get("service_date", "")
            purchase_status_name = row.get("purchase_status_name", "")
            source_order_id = row.get("source_order_id", "")
            source_service_date = row.get("source_service_date", "")
            source_purchase_status_name = row.get("source_purchase_status_name", "")
            source_status_name = row.get("source_status_name", "")
            source_notice_preview = row.get("source_notice_preview", "")
            can_autofill = row.get("can_autofill", False)

            st.markdown(f"""
            <div class="{card_cls}">
                <div class="preview-title">目前訂單：{order_id}</div>
                <div class="preview-sub">
                    <b>客戶 / 電話：</b>{target_name} / {phone}<br>
                    <b>地址：</b>{address}<br>
                    <b>目前服務日期：</b>{service_date}　
                    <b>目前付款狀態：</b>{purchase_status_name or "-"}
                </div>
                <hr style="margin:10px 0;">
                <div class="preview-sub">
                    <b>來源訂單：</b>{source_order_id or "無"}<br>
                    <b>來源服務日期：</b>{source_service_date or "-"}　
                    <b>來源付款狀態：</b>{source_purchase_status_name or "-"}　
                    <b>來源服務狀態：</b>{source_status_name or "-"}<br>
                    <b>來源備註：</b>{source_notice_preview or "無"}
                </div>
                <div class="preview-sub" style="margin-top:8px;">
                    <b>建議：</b>{"建議執行" if can_autofill else "無可參照來源，請人工確認"}
                </div>
            </div>
            """, unsafe_allow_html=True)

            if checked and order_id:
                selected_ids.append(order_id)

    render_section("可自動回填", can_rows, "can_autofill", True)
    render_section("無可參照來源", no_rows, "no_source", False)

    st.markdown("---")
    sec("4. 執行確認")
    st.metric("目前勾選", len(selected_ids))
    st.caption("執行後會把來源客服備註寫入目標訂單，並把目標訂單服務狀態改為已處理。")
    return selected_ids


st.title("📋 Memo 自動回填系統")

sec("1. 登入")
col_e, col_p, col_env, col_login, col_unlock = st.columns([3.0, 3.0, 1.2, 1.2, 1.2])
with col_e:
    email = st.text_input("Email")
with col_p:
    password = st.text_input("Password", type="password")
with col_env:
    env_option = st.selectbox("環境", ["prod", "dev"], index=0)
with col_login:
    st.markdown("<div style='height:26px'></div>", unsafe_allow_html=True)
    login_clicked = st.button("Login", use_container_width=True)
with col_unlock:
    st.markdown("<div style='height:26px'></div>", unsafe_allow_html=True)
    unlock_clicked = st.button("解除鎖定", use_container_width=True)

memo.set_env(env_option)

if unlock_clicked:
    st.session_state.is_running = False
    st.session_state.logs = []
    st.success("已解除鎖定")

if login_clicked:
    try:
        reset_before_action(clear_preview=True, clear_selection=True)

        if not email or not password:
            st.error("請先輸入 Email / Password")
        else:
            st.session_state.is_running = True
            ui_log("===== 開始登入 =====")
            memo.set_env(env_option)
            memo.set_runtime_credentials(email, password)

            with st.spinner("登入中，請稍候…"):
                memo.login(ui_logger=ui_log)

            st.session_state.is_logged_in = True
            st.session_state.login_identity = email
            ui_log("✅ Login 成功")
            st.success("登入成功，請往下設定查詢條件。")

    except Exception as e:
        st.session_state.is_logged_in = False
        st.session_state.login_identity = ""
        ui_log(f"❌ Login 失敗：{e}")
        st.error(f"登入失敗：{e}")
    finally:
        st.session_state.is_running = False

if st.session_state.is_logged_in:
    st.caption(f"✅ 已登入：{st.session_state.login_identity}")
else:
    st.info("請先登入後再查詢或執行。")

st.markdown("---")

sec("2. 設定查詢條件")
mode = st.radio(
    "",
    ["By Google Sheet", "By 電話", "By 搜尋條件"],
    horizontal=True,
    label_visibility="collapsed",
)

reset_mode_state_if_changed(mode)

row_spec = ""
force = False
sheet_run_mode = "指定列號"
sheet_limit = 5

phone_text = ""
date_mode = "服務日期"
purchase_status_name = "全部"
start_date = None
end_date = None

sheet_summary_btn = False
search_btn = False
execute_btn = False

if mode == "By Google Sheet":
    sheet_run_mode = st.radio("處理方式", ["指定列號", "依剩餘筆數處理"], horizontal=True)

    if sheet_run_mode == "指定列號":
        c1, c2 = st.columns([5, 1])
        with c1:
            row_spec = st.text_input("列號（例：2,3,5-8）")
        with c2:
            st.markdown("<div style='height:26px'></div>", unsafe_allow_html=True)
            force = st.checkbox("強制重跑")
        execute_btn = st.button("🚀 執行", use_container_width=True, disabled=not st.session_state.is_logged_in)
    else:
        c1, c2 = st.columns(2)
        with c1:
            sheet_summary_btn = st.button("🔍 查詢目前筆數", use_container_width=True, disabled=not st.session_state.is_logged_in)
        with c2:
            sheet_limit = st.number_input("本次處理筆數", min_value=1, value=5)

        if st.session_state.sheet_summary:
            s = st.session_state.sheet_summary
            m1, m2, m3 = st.columns(3)
            m1.metric("總筆數", s.get("total_rows", 0))
            m2.metric("未處理筆數", s.get("pending_rows", 0))
            m3.metric("已處理筆數", s.get("done_rows", 0))

        execute_btn = st.button("🚀 執行前 N 筆未處理資料", use_container_width=True, disabled=not st.session_state.is_logged_in)

elif mode == "By 電話":
    phone_text = st.text_input("電話號碼（可輸入多筆，用逗號分隔）", placeholder="例：0912345678,0922345678")
    st.caption("會先找出『目標訂單』，再比對最近一筆可參照的來源訂單。")

    c1, c2 = st.columns(2)
    with c1:
        search_btn = st.button("🔍 查詢列表", use_container_width=True, disabled=not st.session_state.is_logged_in)
    with c2:
        execute_btn = st.button("🚀 執行勾選項目", use_container_width=True, disabled=not st.session_state.is_logged_in)

else:
    c1, c2 = st.columns(2)
    with c1:
        date_mode = st.selectbox("日期條件", ["服務日期", "購買日期"])
    with c2:
        purchase_status_name = st.selectbox("付款狀態", ["全部", "已付款", "未付款"], index=0)

    c3, c4 = st.columns(2)
    with c3:
        start_date = st.date_input("開始日期", value=None)
    with c4:
        end_date = st.date_input("結束日期", value=None)

    st.caption("搜尋條件固定只撈服務狀態＝未處理的目標訂單，再比對最近的可參照來源。")

    c5, c6 = st.columns(2)
    with c5:
        search_btn = st.button("🔍 查詢列表", use_container_width=True, disabled=not st.session_state.is_logged_in)
    with c6:
        execute_btn = st.button("🚀 執行勾選項目", use_container_width=True, disabled=not st.session_state.is_logged_in)

with st.expander("4. 執行過程", expanded=True):
    log_box = st.empty()
    log_box.code("\n".join(st.session_state.logs[-3000:]) if st.session_state.logs else "尚未執行")

result_container = st.container()
if st.session_state.result is not None:
    render_result(st.session_state.result)

if sheet_summary_btn:
    try:
        st.session_state.is_running = True
        reset_before_action(clear_preview=True, clear_selection=True)
        ui_log("===== 查詢目前筆數 =====")
        with st.spinner("查詢中，請稍候…"):
            st.session_state.sheet_summary = memo.get_sheet_summary(ui_logger=ui_log)
        ui_log("✅ 查詢完成")
    except Exception as e:
        ui_log(f"❌ 查詢失敗：{e}")
        st.error(str(e))
    finally:
        st.session_state.is_running = False

if search_btn:
    try:
        st.session_state.is_running = True
        reset_before_action(clear_preview=True, clear_selection=True)
        ui_log("===== 開始查詢 =====")

        with st.spinner("查詢中，請稍候…"):
            if mode == "By 電話":
                if not phone_text.strip():
                    raise ValueError("請輸入至少一支電話")
                preview_rows = memo.preview_by_phone_multi(
                    phone_text=phone_text.strip(),
                    ui_logger=ui_log
                )
            else:
                start_text = start_date.strftime("%Y/%m/%d") if start_date else ""
                end_text = end_date.strftime("%Y/%m/%d") if end_date else ""
                preview_rows = memo.preview_by_conditions(
                    date_mode=date_mode,
                    date_start=start_text,
                    date_end=end_text,
                    purchase_status_name=purchase_status_name,
                    ui_logger=ui_log,
                )

        st.session_state.preview_rows = preview_rows or []
        ui_log(f"✅ 查詢完成，共 {len(st.session_state.preview_rows)} 筆")
        st.rerun()

    except Exception as e:
        ui_log(f"❌ 查詢錯誤：{e}")
        st.error(str(e))
    finally:
        st.session_state.is_running = False

selected_ids = []
if mode in ["By 電話", "By 搜尋條件"] and st.session_state.preview_rows:
    st.markdown("---")
    selected_ids = render_preview_blocks(st.session_state.preview_rows)

if execute_btn:
    try:
        st.session_state.is_running = True
        reset_before_execute_keep_preview()

        if mode == "By Google Sheet":
            ui_log("===== 開始執行 =====")
            with st.spinner("執行中，請稍候…"):
                if sheet_run_mode == "指定列號":
                    result = memo.main(row_spec=row_spec, force=force, ui_logger=ui_log)
                else:
                    result = memo.main_first_n_pending(limit=int(sheet_limit), ui_logger=ui_log)
        else:
            if not st.session_state.preview_rows:
                raise RuntimeError("請先查詢列表")

            current_selected_ids = []
            for row in st.session_state.preview_rows:
                oid = str(safe_get(row, "order_id", default="")).strip()
                if oid and st.session_state.get(f"pick_{oid}", False):
                    current_selected_ids.append(oid)

            if not current_selected_ids:
                raise RuntimeError("請先勾選要執行的資料")

            ui_log("===== 開始執行勾選項目 =====")
            ui_log(f"勾選筆數：{len(current_selected_ids)}")

            with st.spinner("執行中，請稍候…"):
                result = memo.main_by_selected_order_ids(order_ids=current_selected_ids, ui_logger=ui_log)

        ui_log("===== 執行完成 =====")
        st.session_state.result = result
        render_result(result)

    except Exception as e:
        ui_log(f"❌ 執行錯誤：{e}")
        st.session_state.result = {**DEFAULT_RESULT, "failed": 1, "errors": [str(e)]}
        render_result(st.session_state.result)
    finally:
        st.session_state.is_running = False
