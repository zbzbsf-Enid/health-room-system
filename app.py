import streamlit as st
import pandas as pd
from datetime import datetime, date

st.set_page_config(page_title="衛保組藥品管理系統 (雲端 24H 版)", layout="wide")

# ⚠️ 請只填入您的「試算表 ID」（開頭那一長串亂碼）
SPREADSHEET_ID = "https://docs.google.com/spreadsheets/d/1fqR5nvOGTOnKljryhMwfbAUAvZo5L11Jtsm823Hf8hU/edit?gid=0#gid=0"

# 免 API 金鑰直接讀取 Google 試算表的公開 CSV 連結
def load_data(sheet_name):
    url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet={sheet_name}"
    try:
        return pd.read_csv(url)
    except Exception:
        return pd.DataFrame()

st.title("💊 衛保組藥品管理系統 (雲端 24H 版)")

tab1, tab2, tab3 = st.tabs([
    "⚡ 多品項快速扣庫/發藥", 
    "✏️ 線上庫存異動與新藥入庫", 
    "📦 庫存總覽與資料匯出"
])

# ---------------------------------------------------------
# Tab 1: 領藥與扣庫 (線上試算)
# ---------------------------------------------------------
with tab1:
    meds_df = load_data("medications")
    
    if not meds_df.empty and 'stock' in meds_df.columns:
        # 過濾庫存大於 0 的品項
        meds_df['stock'] = pd.to_numeric(meds_df['stock'], errors='coerce').fillna(0).astype(int)
        valid_meds = meds_df[meds_df['stock'] > 0]
        
        if not valid_meds.empty:
            st.subheader("1. 選擇發藥品項")
            options = [
                f"{r['name']} ({r.get('chinese_name', '')}) | 批號:{r.get('lot_no', 'DEFAULT')} | 剩餘庫存:{r['stock']}" 
                for _, r in valid_meds.iterrows()
            ]
            
            selected_str = st.selectbox("選擇藥品", options)
            selected_row = valid_meds.iloc[options.index(selected_str)]
            
            qty = st.number_input("扣除數量", min_value=1, max_value=int(selected_row['stock']), value=1)
            
            if st.button("✅ 完成發藥並更新畫面", type="primary"):
                # 計算剩餘庫存
                new_stock = int(selected_row['stock']) - qty
                st.success(f"🎉 發藥成功！【{selected_row['name']}】扣除 {qty} 顆/件，剩餘數量：{new_stock}")
                st.info("💡 提示：請至 Tab 3 匯出最新報表，或同步更新至您的總表。")
        else:
            st.warning("目前試算表內無可用藥品庫存。")
    else:
        st.error("無法讀取試算表，請確認已執行【檔案 ➜ 分享 ➜ 發佈到網路】，且 SPREADSHEET_ID 填寫正確。")

# ---------------------------------------------------------
# Tab 2: 線上異動預覽
# ---------------------------------------------------------
with tab2:
    st.subheader("➕ 藥品庫存登記備忘錄")
    with st.form("add_form"):
        name = st.text_input("藥品名稱")
        qty = st.number_input("數量", value=100)
        expiry = st.date_input("有效期限", value=date(2028, 12, 31))
        submit = st.form_submit_button("登記新增項目")
        if submit and name:
            st.success(f"已登記：{name} x {qty}，效期：{expiry}（請定期整理至雲端總表）")

# ---------------------------------------------------------
# Tab 3: 庫存總覽與即時下載
# ---------------------------------------------------------
with tab3:
    st.subheader("📦 目前 Google 雲端藥品庫存總覽")
    meds_data = load_data("medications")
    st.dataframe(meds_data, use_container_width=True)
