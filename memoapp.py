import sys
import streamlit as st
import memo
import accounts

st.set_page_config(page_title="Memo System", layout="wide")

st.title("Memo 自動補備註系統")

st.subheader("環境")
st.write(f"目前環境：`{memo.ENV_NAME}`")
st.write(f"目前網址：`{memo.BASE_URL}`")

st.subheader("處理列")
row_spec = st.text_input("要處理的列", "2,3,5-8")
force = st.checkbox("強制重跑（忽略 V 欄已有值）", value=False)

st.subheader("登入帳密")

with st.expander("台北帳號", expanded=True):
    taipei_email = st.text_input(
        "台北 Email",
        value=accounts.ACCOUNTS.get("台北", {}).get("email", ""),
        key="taipei_email",
    )
    taipei_password = st.text_input(
        "台北 Password",
        value=accounts.ACCOUNTS.get("台北", {}).get("password", ""),
        type="password",
        key="taipei_password",
    )

with st.expander("台中帳號", expanded=False):
    taichung_email = st.text_input(
        "台中 Email",
        value=accounts.ACCOUNTS.get("台中", {}).get("email", ""),
        key="taichung_email",
    )
    taichung_password = st.text_input(
        "台中 Password",
        value=accounts.ACCOUNTS.get("台中", {}).get("password", ""),
        type="password",
        key="taichung_password",
    )

run_btn = st.button("開始執行", type="primary")

if run_btn:
    # 把 UI 輸入值覆寫回 accounts
    if "台北" in accounts.ACCOUNTS:
        accounts.ACCOUNTS["台北"]["email"] = taipei_email
        accounts.ACCOUNTS["台北"]["password"] = taipei_password

    if "台中" in accounts.ACCOUNTS:
        accounts.ACCOUNTS["台中"]["email"] = taichung_email
        accounts.ACCOUNTS["台中"]["password"] = taichung_password

    argv = []
    if force:
        argv.append("--force")
    if row_spec.strip():
        argv.append(row_spec.strip())

    sys.argv = ["memo.py"] + argv

    st.info("執行中...")

    try:
        result = memo.main()

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("執行筆數", result.get("processed", 0))
        col2.metric("完成筆數", result.get("success", 0))
        col3.metric("失敗筆數", result.get("failed", 0))
        col4.metric("略過筆數", result.get("skipped", 0))

        if result.get("errors"):
            st.subheader("失敗明細")
            for err in result["errors"]:
                st.error(err)

        st.success("執行完成")

    except Exception as e:
        st.error(f"執行失敗：{e}")
