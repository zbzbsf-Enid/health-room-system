import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection
from datetime import datetime, date

st.set_page_config(page_title="衛保組藥品管理系統 (雲端 24H 版)", layout="wide")

# ⚠️ 請把下方網址替換為您的 Google 試算表完整網址
GSHEET_URL = "https://docs.google.com/spreadsheets/d/1fqR5nvOGTOnKljryhMwfbAUAvZo5L11Jtsm823Hf8hU/edit?usp=sharing"

# 建立 Google Sheets 連線
conn = st.connection("gsheets", type=GSheetsConnection)

def load_meds():
    return conn.read(spreadsheet=GSHEET_URL, worksheet="medications", ttl=0)

def load_logs():
    return conn.read(spreadsheet=GSHEET_URL, worksheet="logs", ttl=0)

st.title("💊 衛保組藥品管理系統 (雲端 24H 版)")

tab1, tab2, tab3 = st.tabs([
    "⚡ 多品項快速扣庫/發藥", 
    "✏️ 線上庫存異動與新藥入庫", 
    "📦 庫存總覽"
])

# ---------------------------------------------------------
# Tab 1: 領藥與扣庫
# ---------------------------------------------------------
with tab1:
    meds_df = load_meds()
    # 僅顯示庫存大於 0 的品項
    valid_meds = meds_df[meds_df['stock'].fillna(0).astype(int) > 0]
    
    if not valid_meds.empty:
        st.subheader("1. 選擇藥品並加入清單")
        options = [
            f"{r['name']} ({r['chinese_name']}) | 批號:{r['lot_no']} | 效期:{r['expiry']} | 剩餘庫存:{r['stock']}" 
            for _, r in valid_meds.iterrows()
        ]
        
        selected_str = st.selectbox("選擇藥品", options)
        selected_row = valid_meds.iloc[options.index(selected_str)]
        
        qty = st.number_input("扣除數量", min_value=1, max_value=int(selected_row['stock']), value=1)
        
        if st.button("✅ 確定扣庫發藥", type="primary"):
            # 更新庫存
            meds_df.loc[meds_df['name'] == selected_row['name'], 'stock'] = int(selected_row['stock']) - qty
            conn.update(spreadsheet=GSHEET_URL, worksheet="medications", data=meds_df)
            
            # 寫入紀錄
            logs_df = load_logs()
            new_log = pd.DataFrame([{
                'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M"),
                'med_name': selected_row['name'],
                'lot_no': selected_row['lot_no'],
                'qty': qty
            }])
            updated_logs = pd.concat([logs_df, new_log], ignore_index=True)
            conn.update(spreadsheet=GSHEET_URL, worksheet="logs", data=updated_logs)
            
            st.success(f"🎉 完成扣庫！{selected_row['name']} 扣除 {qty} 顆/件")
            st.rerun()
    else:
        st.warning("目前尚無可用庫存資料。")

# ---------------------------------------------------------
# Tab 2: 庫存校正與新藥入庫
# ---------------------------------------------------------
with tab2:
    st.subheader("➕ 新增 / 調整藥品庫存")
    meds_df = load_meds()
    
    with st.form("add_med_form"):
        name = st.text_input("藥品名稱 (英文/商品名)")
        chi_name = st.text_input("中文藥名")
        lot_no = st.text_input("批號", value="DEFAULT")
        stock = st.number_input("數量", min_value=0, value=100)
        expiry = st.date_input("有效期限", value=date(2028, 12, 31))
        notes = st.text_input("用途/備註")
        
        submit = st.form_submit_button("💾 儲存並更新至雲端試算表")
        
        if submit and name:
            new_data = pd.DataFrame([{
                'name': name,
                'chinese_name': chi_name,
                'lot_no': lot_no,
                'stock': stock,
                'expiry': str(expiry),
                'notes': notes,
                'last_updated': str(date.today())
            }])
            updated_meds = pd.concat([meds_df, new_data], ignore_index=True)
            conn.update(spreadsheet=GSHEET_URL, worksheet="medications", data=updated_meds)
            st.success(f"🎉 成功更新【{name}】至雲端庫存！")
            st.rerun()

# ---------------------------------------------------------
# Tab 3: 庫存總覽
# ---------------------------------------------------------
with tab3:
    st.subheader("📦 目前雲端藥品庫存總覽")
    st.dataframe(load_meds(), use_container_width=True)
