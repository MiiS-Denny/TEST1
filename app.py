import streamlit as st
from io import BytesIO
from datetime import datetime
from openpyxl import load_workbook
import streamlit_authenticator as stauth

st.set_page_config(page_title="Excel 上傳→修改→下載（含登入）", page_icon="🔐", layout="centered")

# ====== 登入設定（示範用，請改成你自己的雜湊）======
# 產生雜湊方式：見文末「產生雜湊」小工具
names = ["Alice", "Bob"]
usernames = ["alice", "bob"]

# 產生雜湊（只在本機先跑一次拿結果；正式版請直接貼雜湊字串）
hashed_pw = stauth.Hasher(["Pass123!", "Pass456!"]).generate()

credentials = {
    "usernames": {
        "alice": {"name": "Alice", "email": "alice@demo.com", "password": hashed_pw[0]},
        "bob":   {"name": "Bob",   "email": "bob@demo.com",   "password": hashed_pw[1]},
    }
}

authenticator = stauth.Authenticate(
    credentials,
    "xl_app_cookie",                 # cookie 名稱
    "super_secret_key_change_me",    # cookie 密鑰
    1                                # cookie 有效天數
)

name, auth_status, username = authenticator.login("登入", "main")
if auth_status is False:
    st.error("帳號或密碼錯誤")
elif auth_status is None:
    st.info("請輸入帳密")
else:
    # ====== 通過登入，顯示主功能 ======
    authenticator.logout("登出", "sidebar")
    st.success(f"歡迎，{name}！")
    st.title("📄 Excel 雲端修改器（上傳→修改→下載）")

    uploaded = st.file_uploader("請上傳 .xlsx", type=["xlsx"])
    st.caption("＊檔案僅在記憶體處理，不會長期存放伺服器。")

    with st.expander("選填：要追加的一列資料"):
        date_str = st.text_input("日期（字串，如 20250819）", "")
        c1,c2,c3 = st.columns(3)
        with c1:
            v1 = st.number_input("Value_1", value=0.0, step=1.0, format="%.3f")
            v2 = st.number_input("Value_2", value=0.0, step=1.0, format="%.3f")
        with c2:
            v3 = st.number_input("Value_3", value=0.0, step=1.0, format="%.3f")
            v4 = st.number_input("Value_4", value=0.0, step=1.0, format="%.3f")
        with c3:
            v5 = st.number_input("Value_5", value=0.0, step=1.0, format="%.3f")
            v6 = st.number_input("Value_6", value=0.0, step=1.0, format="%.3f")
        note = st.text_input("備註", "")

    target_sheet = st.text_input("要寫入的工作表（預設 Data）", value="Data")
    add_timestamp = st.checkbox("下載檔名加上時間戳", value=True)

    if uploaded is not None and st.button("開始修改並提供下載"):
        # 讀入並保持公式/格式/圖表
        data = uploaded.read()
        wb = load_workbook(BytesIO(data), data_only=False, keep_vba=False)
        ws = wb[target_sheet] if target_sheet in wb.sheetnames else wb.create_sheet(title=target_sheet)

        # 追加一列
        if date_str:
            if not (len(date_str) == 8 and date_str.isdigit()):
                st.error("日期需為 8 位數字（YYYYMMDD）。")
                st.stop()
            if ws.max_row == 1 and all((cell.value is None) for cell in ws[1]):
                ws.append(["date_str","value_1","value_2","value_3","value_4","value_5","value_6","note"])
            ws.append([date_str, v1, v2, v3, v4, v5, v6, note])
            # 第一欄強制文字格式（避免被 Excel 轉日期）
            for cell in ws["A"]:
                cell.number_format = "@"

        out = BytesIO()
        wb.save(out)
        out.seek(0)

        base = uploaded.name.rsplit(".xlsx", 1)[0]
        fname = f"{base}-{datetime.now().strftime('%Y%m%d-%H%M%S')}.xlsx" if add_timestamp else f"{base}.xlsx"

        st.download_button(
            "📥 下載修改後的 Excel",
            data=out.getvalue(),
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.info("若圖表來源綁『表格 (Ctrl+T)』或動態範圍，追加資料後打開檔案圖表會自動延伸。")

