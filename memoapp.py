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
    padding-top: 1.2rem !important;
    padding-bottom: 1rem !important;
    max-width: 1220px !important;
}
h1 {
    font-size: 20px !important;
    font-weight: 700 !important;
    margin: 0 0 16px 0 !important;
}
.sec-label {
    font-size: 11px;
    font-weight: 700;
    color: #999;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    margin: 4px 0 8px 0;
}
[data-testid="stButton"] > button {
    background: #111 !important;
    color: #fff !important;
    border: none !important;
    border-radius: 8px !important;
    font-size: 14px !important;
    font-weight: 600 !important;
    padding: 9px 0 !important;
}
[data-testid="stButton"] > button:hover {
    background: #333 !important;
}
[data-testid="stButton"] > button:disabled {
    background: #ccc !important;
}
[data-testid="stCode"] {
    font-size: 11.5px !important;
    background: #13131f !important;
    border-radius: 8px !important;
}
[data-testid="stMetric"] {
    background: #fff !important;
    border: 1px solid #e6e8ec !important;
    border-radius: 10px !important;
    padding: 12px !important;
}
hr {
    border-color: #ececec !important;
    margin: 14px 0 !important;
}
.small-muted {
    color: #888;
    font-size: 12px;
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

    "address_member_candidates": [],
    "address_selected_email": "",
    "address_list": [],
    "address_selected_value": "",
    "address_preview": None,
}

for k, v in DEFAULT_STATE.items():
    if k not in st.session_state:
        st.session_state[k] = v


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
        st.markdown("<hr>", unsafe_allow_html=True)
        sec("6. 執行結果")
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


def clear_address_flow_state():
    st.session_state.address_member_candidates = []
    st.session_state.address_selected_email = ""
    st.session_state.address_list = []
    st.session_state.address_selected_value = ""
    st.session_state.address_preview = None


def reset_before_action(clear_preview=True, clear_selection=True):
    st.session_state.logs = []
    st.session_state.result = None

    if clear_preview:
        st.session_state.preview_rows = []
        st.session_state.sheet_summary = None

    if clear_selection:
        clear_pick_states()

    clear_address_flow_state()

    try:
        log_box.code("尚未執行")
    except Exception:
        pass


def reset_before_execute_keep_preview():
    st.session_state.logs = []
    st.session_state.result = None
    clear_address_flow_state()
    try:
        log_box.code("尚未執行")
    except Exception:
        pass


def reset_mode_state_if_changed(current_mode):
    if st.session_state.last_mode != current_mode:
        st.session_state.preview_rows = []
        st.session_state.sheet_summary = None
        clear_pick_states()
        clear_address_flow_state()
        st.session_state.last_mode = current_mode


def render_result_preview_blocks(rows):
    sec("3. 查詢結果預覽")

    if not rows:
        st.info("查無資料")
        return []

    total_count = len(rows)
    same_count = sum(1 for r in rows if r.get("has_same_address_history"))
    diff_count = total_count - same_count
    paid_count = sum(1 for r in rows if r.get("purchase_status_name") == "已付款")
    unpaid_count = sum(1 for r in rows if r.get("purchase_status_name") == "未付款")

    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("查詢總筆數", total_count)
    m2.metric("有同地址", same_count)
    m3.metric("無同地址", diff_count)
    m4.metric("已付款", paid_count)
    m5.metric("未付款", unpaid_count)

    st.info("請先勾選要處理的資料，再按下方「執行勾選項目」。")

    grouped = {
        "same_paid": [],
        "same_unpaid": [],
        "diff_paid": [],
        "diff_unpaid": [],
    }

    for row in rows:
        has_same = bool(row.get("has_same_address_history"))
        paid = row.get("purchase_status_name") == "已付款"

        if has_same and paid:
            grouped["same_paid"].append(row)
        elif has_same and not paid:
            grouped["same_unpaid"].append(row)
        elif (not has_same) and paid:
            grouped["diff_paid"].append(row)
        else:
            grouped["diff_unpaid"].append(row)

    selected_ids = []

    def render_block(title, block_rows, block_key):
        st.markdown(f"#### {title}")
        if not block_rows:
            st.caption("沒有資料")
            return

        c1, c2, c3 = st.columns([1, 1, 4])
        with c1:
            if st.button("本區全選", key=f"btn_sel_{block_key}", use_container_width=True):
                for row in block_rows:
                    oid = str(safe_get(row, "order_id", default="")).strip()
                    if oid:
                        st.session_state[f"pick_{oid}"] = True
                st.rerun()
        with c2:
            if st.button("本區全不選", key=f"btn_unsel_{block_key}", use_container_width=True):
                for row in block_rows:
                    oid = str(safe_get(row, "order_id", default="")).strip()
                    if oid:
                        st.session_state[f"pick_{oid}"] = False
                st.rerun()
        with c3:
            st.caption(f"本區共 {len(block_rows)} 筆")

        header = st.columns([0.8, 1.4, 1.6, 2.8, 1.2, 1.0, 1.0, 1.0, 1.0, 1.2])
        names = ["選取", "訂單編號", "姓名 / 電話", "地址", "日期", "同地址", "不同地址", "付款", "同址", "建議"]
        for col, name in zip(header, names):
            col.markdown(f"**{name}**")

        for row in block_rows:
            order_id = str(safe_get(row, "order_id", default="")).strip()
            customer_name = safe_get(row, "customer_name", "name", default="")
            phone = safe_get(row, "phone", default="")
            address = safe_get(row, "address", default="")
            service_date = safe_get(row, "service_date", default="")
            same_addr_count = int(safe_get(row, "same_address_count", default=0) or 0)
            diff_addr_count = int(safe_get(row, "different_address_count", default=0) or 0)
            purchase_status_name = safe_get(row, "purchase_status_name", default="")
            has_same = bool(safe_get(row, "has_same_address_history", default=False))

            cols = st.columns([0.8, 1.4, 1.6, 2.8, 1.2, 1.0, 1.0, 1.0, 1.0, 1.2])

            default_checked = has_same
            checked = cols[0].checkbox(
                "選取",
                key=f"pick_{order_id}",
                label_visibility="collapsed",
                value=st.session_state.get(f"pick_{order_id}", default_checked),
            )
            cols[1].write(order_id)
            cols[2].write(f"{customer_name}\n{phone}".strip())
            cols[3].write(address)
            cols[4].write(service_date)
            cols[5].write(f"{same_addr_count} 筆")
            cols[6].write(f"{diff_addr_count} 筆")
            cols[7].write(purchase_status_name or "-")
            cols[8].write("有" if has_same else "無")
            cols[9].write("建議優先" if has_same else "人工確認")

            if checked and order_id:
                selected_ids.append(order_id)

    render_block("有同地址紀錄｜已付款", grouped["same_paid"], "same_paid")
    render_block("有同地址紀錄｜未付款", grouped["same_unpaid"], "same_unpaid")
    render_block("無同地址紀錄｜已付款", grouped["diff_paid"], "diff_paid")
    render_block("無同地址紀錄｜未付款", grouped["diff_unpaid"], "diff_unpaid")

    st.markdown("<hr>", unsafe_allow_html=True)
    sec("4. 執行確認")

    selected_rows = [r for r in rows if str(safe_get(r, "order_id", default="")).strip() in selected_ids]
    sel_total = len(selected_rows)
    sel_same = sum(1 for r in selected_rows if r.get("has_same_address_history"))
    sel_diff = sel_total - sel_same
    sel_paid = sum(1 for r in selected_rows if r.get("purchase_status_name") == "已付款")
    sel_unpaid = sum(1 for r in selected_rows if r.get("purchase_status_name") == "未付款")

    x1, x2, x3, x4, x5 = st.columns(5)
    x1.metric("目前勾選", sel_total)
    x2.metric("有同地址", sel_same)
    x3.metric("無同地址", sel_diff)
    x4.metric("已付款", sel_paid)
    x5.metric("未付款", sel_unpaid)

    st.caption("確認後將開始回填客服備註並更新狀態。")
    return selected_ids


st.title("📋 Memo 自動回填系統")

sec("1. 登入")
col_e, col_p, col_env, col_login = st.columns([3.0, 3.0, 1.2, 1.2])
with col_e:
    email = st.text_input("Email")
with col_p:
    password = st.text_input("Password", type="password")
with col_env:
    env_option = st.selectbox("環境", ["prod", "dev"], index=0)
with col_login:
    st.markdown("<div style='height:26px'></div>", unsafe_allow_html=True)
    login_clicked = st.button("Login", use_container_width=True, disabled=st.session_state.is_running)

memo.set_env(env_option)

if login_clicked:
    try:
        reset_before_action(clear_preview=True, clear_selection=True)

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

st.markdown("<hr>", unsafe_allow_html=True)

sec("2. 設定查詢條件")
mode = st.radio(
    "",
    ["By Google Sheet", "By 電話", "By 搜尋條件", "By 地址備註更新"],
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
limit = 5
start_date = None
end_date = None

address_phone = ""
address_mail = ""
base_service_date = None
new_member_notice = ""

sheet_summary_btn = False
search_btn = False
execute_btn = False
address_find_member_btn = False
address_preview_btn = False
address_execute_btn = False

if mode == "By Google Sheet":
    sheet_run_mode = st.radio("處理方式", ["指定列號", "依剩餘筆數處理"], horizontal=True)

    if sheet_run_mode == "指定列號":
        c1, c2 = st.columns([5, 1])
        with c1:
            row_spec = st.text_input("列號（例：2,3,5-8）")
        with c2:
            st.markdown("<div style='height:26px'></div>", unsafe_allow_html=True)
            force = st.checkbox("強制重跑")
        execute_btn = st.button(
            "🚀 執行",
            use_container_width=True,
            disabled=(st.session_state.is_running or not st.session_state.is_logged_in),
        )
    else:
        c1, c2 = st.columns(2)
        with c1:
            sheet_summary_btn = st.button(
                "🔍 查詢目前筆數",
                use_container_width=True,
                disabled=(st.session_state.is_running or not st.session_state.is_logged_in),
            )
        with c2:
            sheet_limit = st.number_input("本次處理筆數", min_value=1, value=5)

        if st.session_state.sheet_summary:
            s = st.session_state.sheet_summary
            m1, m2, m3 = st.columns(3)
            m1.metric("總筆數", s.get("total_rows", 0))
            m2.metric("未處理筆數", s.get("pending_rows", 0))
            m3.metric("已處理筆數", s.get("done_rows", 0))

        execute_btn = st.button(
            "🚀 執行前 N 筆未處理資料",
            use_container_width=True,
            disabled=(st.session_state.is_running or not st.session_state.is_logged_in),
        )

elif mode == "By 電話":
    phone_text = st.text_input("電話號碼（可輸入多筆，用逗號分隔）", placeholder="例：0912345678,0922345678")
    st.caption("可輸入多筆電話，用逗號分隔；會先查詢列表，再勾選執行。")

    c1, c2 = st.columns(2)
    with c1:
        search_btn = st.button(
            "🔍 查詢列表",
            use_container_width=True,
            disabled=(st.session_state.is_running or not st.session_state.is_logged_in),
        )
    with c2:
        execute_btn = st.button(
            "🚀 執行勾選項目",
            use_container_width=True,
            disabled=(st.session_state.is_running or not st.session_state.is_logged_in),
        )

elif mode == "By 搜尋條件":
    c1, c2, c3 = st.columns([1.4, 1.4, 1.0])
    with c1:
        date_mode = st.selectbox("日期條件", ["服務日期", "購買日期"])
    with c2:
        purchase_status_name = st.selectbox("付款狀態", ["全部", "已付款", "未付款"], index=0)
    with c3:
        limit = st.number_input("處理筆數", min_value=1, value=5)

    c4, c5 = st.columns(2)
    with c4:
        start_date = st.date_input("開始日期", value=None)
    with c5:
        end_date = st.date_input("結束日期", value=None)

    c6, c7 = st.columns(2)
    with c6:
        search_btn = st.button(
            "🔍 查詢列表",
            use_container_width=True,
            disabled=(st.session_state.is_running or not st.session_state.is_logged_in),
        )
    with c7:
        execute_btn = st.button(
            "🚀 執行勾選項目",
            use_container_width=True,
            disabled=(st.session_state.is_running or not st.session_state.is_logged_in),
        )

else:
    c1, c2, c3 = st.columns([2, 2, 1.2])
    with c1:
        address_phone = st.text_input("電話（優先）")
    with c2:
        address_mail = st.text_input("Mail（電話無法唯一對應時再補）")
    with c3:
        base_service_date = st.date_input("基準服務日期", value=None)

    address_find_member_btn = st.button(
        "🔍 查詢會員 / 地址",
        use_container_width=True,
        disabled=(st.session_state.is_running or not st.session_state.is_logged_in),
    )

    if st.session_state.address_member_candidates:
        sec("3. 會員確認")
        if len(st.session_state.address_member_candidates) == 1:
            st.session_state.address_selected_email = st.session_state.address_member_candidates[0]
            st.success(f"已對應會員：{st.session_state.address_selected_email}")
        else:
            st.warning("此電話無法唯一對應會員，請選擇或改輸入 mail。")
            st.session_state.address_selected_email = st.selectbox(
                "選擇會員 mail",
                st.session_state.address_member_candidates,
                index=0,
            )

    if st.session_state.address_list:
        sec("4. 選擇要更新的地址")
        address_options = [x["address"] for x in st.session_state.address_list]
        st.session_state.address_selected_value = st.radio(
            "地址清單",
            options=address_options,
            index=0 if address_options else None,
        )

        selected_addr_obj = next(
            (x for x in st.session_state.address_list if x["address"] == st.session_state.address_selected_value),
            None
        )
        if selected_addr_obj:
            st.caption(f"目前備註摘要：{selected_addr_obj.get('memo_summary', '') or '(空)'}")

        new_member_notice = st.text_area("新客服備註", height=180)

        c4, c5 = st.columns(2)
        with c4:
            address_preview_btn = st.button(
                "🔍 預估更新筆數",
                use_container_width=True,
                disabled=(st.session_state.is_running or not st.session_state.is_logged_in),
            )
        with c5:
            address_execute_btn = st.button(
                "🚀 確認並更新",
                use_container_width=True,
                disabled=(st.session_state.is_running or not st.session_state.is_logged_in),
            )

with st.expander("5. 執行過程", expanded=True):
    log_box = st.empty()
    log_box.code("\n".join(st.session_state.logs[-3000:]) if st.session_state.logs else "尚未執行")

selected_ids = []
if mode in ["By 電話", "By 搜尋條件"] and st.session_state.preview_rows:
    st.markdown("<hr>", unsafe_allow_html=True)
    selected_ids = render_result_preview_blocks(st.session_state.preview_rows)

if mode == "By 地址備註更新" and st.session_state.address_preview:
    st.markdown("<hr>", unsafe_allow_html=True)
    sec("5. 預估更新結果")
    preview = st.session_state.address_preview
    c1, c2, c3 = st.columns(3)
    c1.metric("會員 mail", preview.get("email", ""))
    c2.metric("地址", preview.get("address", ""))
    c3.metric("預計更新筆數", preview.get("count", 0))

    if preview.get("items"):
        st.markdown("#### 預計更新清單")
        for item in preview["items"]:
            st.write(f"{item.get('service_date','')}｜{item.get('order_no','')}｜{item.get('name','')}｜{item.get('address','')}")

result_container = st.container()
if st.session_state.result is not None:
    render_result(st.session_state.result)

if sheet_summary_btn:
    if not st.session_state.is_logged_in:
        st.warning("請先登入")
        st.stop()
    try:
        st.session_state.is_running = True
        reset_before_action(clear_preview=True, clear_selection=True)
        ui_log("===== 查詢目前筆數 =====")

        with st.spinner("查詢中，請稍候…"):
            st.session_state.sheet_summary = memo.get_sheet_summary(ui_logger=ui_log)

        ui_log("✅ 查詢完成")
        st.rerun()

    except Exception as e:
        ui_log(f"❌ 查詢失敗：{e}")
        st.error(str(e))
    finally:
        st.session_state.is_running = False

if search_btn:
    if not st.session_state.is_logged_in:
        st.warning("請先登入")
        st.stop()

    try:
        st.session_state.is_running = True
        reset_before_action(clear_preview=True, clear_selection=True)
        ui_log("===== 開始查詢 =====")

        with st.spinner("查詢中，請稍候…"):
            if mode == "By 電話":
                if not phone_text.strip():
                    raise ValueError("請輸入至少一支電話")
                preview_rows = memo.preview_by_phone_multi(phone_text=phone_text.strip(), ui_logger=ui_log)
            else:
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

        st.session_state.preview_rows = preview_rows or []
        ui_log(f"✅ 查詢完成，共 {len(st.session_state.preview_rows)} 筆")
        st.rerun()

    except Exception as e:
        ui_log(f"❌ 查詢錯誤：{e}")
        st.error(str(e))
    finally:
        st.session_state.is_running = False

if address_find_member_btn:
    if not st.session_state.is_logged_in:
        st.warning("請先登入")
        st.stop()

    try:
        st.session_state.is_running = True
        reset_before_action(clear_preview=True, clear_selection=True)

        if not address_phone.strip() and not address_mail.strip():
            raise ValueError("請至少輸入電話或 mail")
        if not base_service_date:
            raise ValueError("請先選擇基準服務日期")

        ui_log("===== 開始查詢會員 / 地址 =====")
        session = memo.login(ui_logger=ui_log)

        emails = []
        if address_phone.strip():
            emails = memo.find_member_by_phone(session, address_phone.strip())
            if len(emails) == 1:
                st.session_state.address_selected_email = emails[0]

        if address_mail.strip():
            if address_mail.strip() not in emails:
                emails.append(address_mail.strip())
            st.session_state.address_selected_email = address_mail.strip()

        if not emails:
            raise RuntimeError("找不到對應會員，請改輸入 mail")
        st.session_state.address_member_candidates = emails

        selected_email = st.session_state.address_selected_email or emails[0]
        address_rows = memo.find_all_address_memo(session, selected_email)
        if not address_rows:
            raise RuntimeError("找不到此會員的地址備註資料")

        st.session_state.address_list = address_rows

        ui_log(f"✅ 找到會員候選 {len(emails)} 個，地址 {len(address_rows)} 筆")
        st.rerun()

    except Exception as e:
        ui_log(f"❌ 查詢失敗：{e}")
        st.error(str(e))
    finally:
        st.session_state.is_running = False

if address_preview_btn:
    if not st.session_state.is_logged_in:
        st.warning("請先登入")
        st.stop()

    try:
        st.session_state.is_running = True
        st.session_state.logs = []
        st.session_state.result = None
        st.session_state.address_preview = None
        log_box.code("尚未執行")

        if not st.session_state.address_selected_email:
            raise ValueError("請先查詢並確認會員")
        if not st.session_state.address_selected_value:
            raise ValueError("請先選擇地址")
        if not base_service_date:
            raise ValueError("請先選擇基準服務日期")

        ui_log("===== 開始預估更新筆數 =====")
        session = memo.login(ui_logger=ui_log)
        items = memo.preview_future_orders_by_member_and_address(
            session=session,
            keyword=st.session_state.address_selected_email,
            address=st.session_state.address_selected_value,
            service_date=base_service_date.strftime("%Y/%m/%d"),
        )

        st.session_state.address_preview = {
            "email": st.session_state.address_selected_email,
            "address": st.session_state.address_selected_value,
            "count": len(items),
            "items": items,
        }
        ui_log(f"✅ 預計更新 {len(items)} 筆")
        st.rerun()

    except Exception as e:
        ui_log(f"❌ 預估失敗：{e}")
        st.error(str(e))
    finally:
        st.session_state.is_running = False

if execute_btn:
    if not st.session_state.is_logged_in:
        st.warning("請先登入")
        st.stop()

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
                st.warning("請先查詢列表")
                st.stop()

            current_selected_ids = []
            for row in st.session_state.preview_rows:
                oid = str(safe_get(row, "order_id", default="")).strip()
                if oid and st.session_state.get(f"pick_{oid}", False):
                    current_selected_ids.append(oid)

            if not current_selected_ids:
                st.warning("請先勾選要執行的資料")
                st.stop()

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

if address_execute_btn:
    if not st.session_state.is_logged_in:
        st.warning("請先登入")
        st.stop()

    try:
        st.session_state.is_running = True
        st.session_state.logs = []
        st.session_state.result = None
        log_box.code("尚未執行")

        if not st.session_state.address_selected_email:
            raise ValueError("請先查詢並確認會員")
        if not st.session_state.address_selected_value:
            raise ValueError("請先選擇地址")
        if not base_service_date:
            raise ValueError("請先選擇基準服務日期")
        if not new_member_notice.strip():
            raise ValueError("請輸入新客服備註")

        ui_log("===== 開始地址備註批次更新 =====")
        with st.spinner("執行中，請稍候…"):
            result = memo.update_future_orders_by_member_and_address(
                keyword=st.session_state.address_selected_email,
                address=st.session_state.address_selected_value,
                service_date=base_service_date.strftime("%Y/%m/%d"),
                new_notice=new_member_notice.strip(),
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
