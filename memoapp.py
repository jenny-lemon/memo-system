# app.py
import streamlit as st
import sys
import memo

st.set_page_config(page_title="Memo System", layout="wide")

st.title("📋 Memo 自動補備註系統")

st.write("輸入要處理的列，例如：")
st.code("2,3,5-8")

row_spec = st.text_input("處理列", "2,3,5-8")

force = st.checkbox("強制重跑（忽略已成功）", value=False)

if st.button("🚀 開始執行"):
    st.info("執行中...")

    # 模擬 CLI 參數
    argv = []
    if force:
        argv.append("--force")
    if row_spec.strip():
        argv.append(row_spec.strip())

    sys.argv = ["memo.py"] + argv

    try:
        memo.main()
        st.success("完成 ✅")
    except Exception as e:
        st.error(f"錯誤：{e}")
