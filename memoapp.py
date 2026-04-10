# -*- coding: utf-8 -*-
import streamlit as st
import memo

# ========================
# 頁面設定
# ========================
st.set_page_config(page_title="Memo 自動回填系統", layout="wide")

st.title("📋 Memo 自動回填系統")

# ========================
# 環境切換
# ========================
st.sidebar.header("⚙️ 環境設定")

env_option = st.sidebar.radio(
    "選擇環境",
    ["prod", "dev"],
    index=0
)

memo.set_env(env_option)

st.sidebar.success(f"目前環境：{env_option}")

# ========================
# 模式選擇
# ========================
mode = st.radio(
    "選擇執行模式",
    [
        "📄 Google Sheet列號",
        "📱 手機號碼",
        "🔍 搜尋條件"
    ]
)

# ========================
# 共用 log 顯示
# ========================
log_placeholder = st.empty()

def ui_logger(msg):
    with log_placeholder.container():
        st.text(msg)

# ========================
# 模式 1：Sheet
# ========================
if mode == "📄 Google Sheet列號":

    st.subheader("📄 依 Google Sheet 列號執行")

    row_spec = st.text_input("輸入列號（例：2,3,5-8）", "2")

    force = st.checkbox("強制覆寫（忽略已有結果）")

    if st.button("🚀 開始執行"):

        log_placeholder.empty()

        result = memo.main(
            row_spec=row_spec,
            force=force,
            ui_logger=ui_logger
        )

        st.success("執行完成")
        st.write(result)


# ========================
# 模式 2：電話
# ========================
elif mode == "📱 手機號碼":

    st.subheader("📱 依電話處理")

    phone = st.text_input("輸入手機號碼")

    if st.button("🚀 開始執行"):

        log_placeholder.empty()

        result = memo.main_by_phone(
            phone=phone,
            ui_logger=ui_logger
        )

        st.success("執行完成")
        st.write(result)


# ========================
# 模式 3：搜尋條件（你新增的）
# ========================
elif mode == "🔍 搜尋條件":

    st.subheader("🔍 條件批次處理")

    col1, col2 = st.columns(2)

    with col1:
        date_s = st.text_input("訂購日期 (YYYY/MM/DD)", "2026/04/09")

    with col2:
        limit = st.number_input("每次處理筆數", min_value=1, max_value=100, value=5)

    st.info("""
條件說明：
- 訂購日期 = 指定日期
- 付款狀態：非取消
- 服務狀態：未處理
- 自動抓前一次「相同電話 + 相同地址 + 已付款 + 已處理」
""")

    if st.button("🚀 開始執行"):

        log_placeholder.empty()

        result = memo.main_by_conditions(
            date_s=date_s,
            limit=int(limit),
            ui_logger=ui_logger
        )

        st.success("執行完成")
        st.write(result)
