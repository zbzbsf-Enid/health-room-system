from monthly_report import generate_monthly_report_excel
import streamlit as st
import pandas as pd
import io
from datetime import datetime, date

st.set_page_config(page_title="衛保組藥品關懷管理系統", layout="wide")

# 🎨 注入溫暖系 CSS 樣式（圓角卡片、柔和配色、放大字體 18px+）
st.markdown("""
    <style>
    /* 全域內文與字體放大 */
    html, body, [class*="st-"], .stMarkdown, p, span {
        font-size: 18px !important;
        font-family: "Microsoft JhengHei", "PingFang TC", sans-serif !important;
    }
    
    /* 標籤與欄位名稱 */
    label, .stSelectbox label, .stNumberInput label, .stTextInput label, .stDateInput label, .stRadio label {
        font-size: 18px !important;
        font-weight: 600 !important;
        color: #d96d00 !important; /* 溫暖暖橘色 */
    }
    
    /* 下拉選單與輸入框圓角與質感 */
    .stSelectbox div[data-baseweb="select"] > div, 
    .stNumberInput input, 
    .stTextInput input {
        font-size: 18px !important;
        border-radius: 10px !important;
    }
    
    /* 分頁頁籤樣式 */
    button[data-baseweb="tab"] {
        font-size: 20px !important;
        font-weight: bold !important;
        padding: 10px 20px !important;
    }
    
    /* 溫暖系按鈕設計 */
    .stButton > button {
        font-size: 18px !important;
        font-weight: bold !important;
        border-radius: 12px !important;
        padding: 0.6rem 1.2rem !important;
        box-shadow: 0 2px 5px rgba(0,0,0,0.1) !important;
        transition: all 0.2s ease !important;
    }
    
    /* 表格樣式優化 */
    .stDataFrame {
        font-size: 16px !important;
        border-radius: 10px !important;
    }
    
    /* 提示訊息框 */
    .stAlert {
        font-size: 18px !important;
        border-radius: 12px !important;
    }
    </style>
""", unsafe_allow_html=True)

# Google 試算表 ID
SPREADSHEET_ID = "1fqR5nvOGTOnKljryhMwfbAUAvZo5L11Jtsm823Hf8hU"

@st.cache_data(ttl=5)
def load_data(sheet_name):
    url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet={sheet_name}"
    try:
        df = pd.read_csv(url)
        df.columns = [str(c).strip() for c in df.columns]
        return df
    except Exception:
        return pd.DataFrame()

# 初始化購物車暫存
if 'cart' not in st.session_state:
    st.session_state.cart = []

# 溫馨系統標題
st.title("🌸 衛保組藥品關懷管理系統")
st.caption("☀️ 守護校園健康每一天，感謝您的細心付出與照護！")

tab1, tab2, tab3 = st.tabs([
    "✨ 多品項快速扣庫與領藥", 
    "🌿 庫存校正與新藥進貨紀錄", 
    "📦 雲端總覽與 Excel 報表匯出"
])

# ---------------------------------------------------------
# Tab 1: 快速扣庫/發藥 (親切語調)
# ---------------------------------------------------------
with tab1:
    meds_df = load_data("medications")
    
    if not meds_df.empty and 'stock' in meds_df.columns:
        meds_df['stock'] = pd.to_numeric(meds_df['stock'], errors='coerce').fillna(0).astype(int)
        valid_meds = meds_df[meds_df['stock'] > 0].copy()
        
        if not valid_meds.empty:
            st.subheader("1. 選擇給藥品項")
            
            options = [
                f"{r.get('name', '')} ({r.get('chinese_name', '')}) | 批號:{r.get('lot_no', 'DEFAULT')} | 效期:{r.get('expiry', '')} | 剩餘庫存:{r['stock']}" 
                for _, r in valid_meds.iterrows()
            ]
            
            col_select, col_qty, col_add = st.columns([3.5, 1.2, 1.3])
            
            with col_select:
                selected_str = st.selectbox("選擇藥品 (已自動為您優先帶出最早到期的批次 💡)", options, key="dispense_select")
                selected_row = valid_meds.iloc[options.index(selected_str)]
                
            with col_qty:
                qty = st.number_input("領取數量", min_value=1, max_value=int(selected_row['stock']), value=1, key="dispense_qty")
                
            with col_add:
                st.write("")
                st.write("")
                if st.button("💖 加入發藥清單", use_container_width=True, type="primary"):
                    existing = next((item for item in st.session_state.cart if item['name'] == selected_row['name'] and item['lot_no'] == selected_row.get('lot_no', 'DEFAULT')), None)
                    if existing:
                        if existing['qty'] + qty <= selected_row['stock']:
                            existing['qty'] += qty
                            st.toast(f"已更新【{selected_row['name']}】數量為 {existing['qty']} 顆/件囉！", icon="✨")
                        else:
                            st.error("發藥數量已超過庫存上限，請確認後重試～")
                    else:
                        st.session_state.cart.append({
                            'name': selected_row['name'],
                            'chinese_name': selected_row.get('chinese_name', ''),
                            'lot_no': selected_row.get('lot_no', 'DEFAULT'),
                            'qty': qty,
                            'max_stock': selected_row['stock']
                        })
                        st.toast(f"已貼心為您加入：{selected_row['name']} x {qty}", icon="🌸")
                    st.rerun()

            st.markdown("---")
            
            cart_count = len(st.session_state.cart)
            st.subheader(f"2. 本次待發藥品清單 (已選取 {cart_count} 項品項)")

            if cart_count > 0:
                cart_df = pd.DataFrame(st.session_state.cart)[['name', 'chinese_name', 'lot_no', 'qty']]
                cart_df.columns = ['商品名', '中文名', '批號', '領用數量']
                
                st.dataframe(cart_df, use_container_width=True, height=200)
                
                btn_col1, btn_col2 = st.columns([1, 3])
                with btn_col1:
                    if st.button("🧹 清空待發清單", use_container_width=True):
                        st.session_state.cart = []
                        st.rerun()
                        
                with btn_col2:
                    if st.button("🎉 完成發藥扣庫紀錄", type="primary", use_container_width=True):
                        summary_list = [f"{item['name']} x{item['qty']}" for item in st.session_state.cart]
                        st.balloons()
                        st.success(f"👏 太棒了！本次發藥試算完成：{', '.join(summary_list)}")
                        st.info("💡 溫馨提醒：請記得至 Google 試算表同步更新扣除後的庫存數量喔！")
                        st.session_state.cart = []
            else:
                st.info("☘️ 目前待發清單是空的～請從上方選擇藥品並點擊【💖 加入發藥清單】")

        else:
            st.warning("目前試算表內尚無可用庫存（庫存皆為 0）。")
    else:
        st.error("連線遇到了一點問題，請確認 Google 試算表已開啟【檔案 ➜ 分享 ➜ 發佈到網路】喔！")

# ---------------------------------------------------------
# Tab 2: 線上庫存校正與新藥進貨
# ---------------------------------------------------------
with tab2:
    st.subheader("🌿 線上庫存校正與新藥進貨管理")
    
    subtab1, subtab2 = st.tabs([
        "🔄 既有藥品 — 盤點校正 / 效期變更 / 補充庫存",
        "➕ 建立全新藥品 / 新批次進貨"
    ])
    
    meds_df = load_data("medications")
    
    # 子分頁 1：盤點與校正
    with subtab1:
        if not meds_df.empty:
            meds_df['stock'] = pd.to_numeric(meds_df['stock'], errors='coerce').fillna(0).astype(int)
            
            options = [
                f"{r.get('name', '')} ({r.get('chinese_name', '')}) | 批號:{r.get('lot_no', 'DEFAULT')} | 庫存:{r['stock']} | 效期:{r.get('expiry', '')}"
                for _, r in meds_df.iterrows()
            ]
            
            selected_str = st.selectbox("🔍 請選擇欲維護/盤點的藥品項目：", options, key="edit_select")
            selected_row = meds_df.iloc[options.index(selected_str)]
            
            st.write("")
            
            # 溫馨數據卡片
            m_col1, m_col2, m_col3 = st.columns(3)
            with m_col1:
                st.caption("目前登記庫存")
                st.markdown(f"## **{selected_row['stock']} 顆/件**")
            with m_col2:
                st.caption("當前有效期限")
                st.markdown(f"## **{selected_row.get('expiry', '未填寫')}**")
            with m_col3:
                st.caption("最後盤點日期")
                last_up = str(selected_row.get('last_updated', '無紀錄'))
                st.markdown(f"## **{last_up if last_up != 'nan' else '無紀錄'}**")
                
            st.markdown("---")
            
            with st.form("edit_form"):
                c1, c2 = st.columns(2)
                with c1:
                    adj_mode = st.radio(
                        "選擇調整模式：",
                        ["直接修正為 (盤點覆蓋)", "舊批次補貨 (數量累加)"]
                    )
                with c2:
                    new_stock_val = st.number_input("盤點後正確數量", min_value=0, value=int(selected_row['stock']))
                    
                c3, c4 = st.columns(2)
                with c3:
                    try:
                        def_exp = datetime.strptime(str(selected_row.get('expiry', '')), "%Y-%m-%d").date()
                    except:
                        def_exp = date(2028, 5, 31)
                    new_exp = st.date_input("校正有效期限", value=def_exp)
                with c4:
                    op_date = st.date_input("異動/盤點日期", value=date.today())
                    
                notes_val = st.text_input("用途 / 備註說明", value=str(selected_row.get('notes', '')) if str(selected_row.get('notes', '')) != 'nan' else "")
                
                submit_edit = st.form_submit_button("🌱 儲存盤點校正紀錄", type="primary", use_container_width=True)
                if submit_edit:
                    st.success(f"已順利登記【{selected_row['name']}】異動紀錄！辛苦了，請記得同步至 Google 試算表～")
        else:
            st.info("目前尚無藥品資料可以校正。")

    # 子分頁 2：新增新藥
    with subtab2:
        st.markdown("##### ➕ 新增全新藥品品項或新批次")
        with st.form("add_new_med_form"):
            col1, col2, col3 = st.columns(3)
            with col1:
                new_name = st.text_input("藥品名稱 (英文/商品名)")
            with col2:
                new_chi = st.text_input("中文藥名")
            with col3:
                new_lot = st.text_input("批號", value="DEFAULT")
                
            col4, col5, col6 = st.columns(3)
            with col4:
                new_stock = st.number_input("進貨數量", min_value=1, value=100)
            with col5:
                new_exp = st.date_input("有效期限", value=date(2028, 12, 31))
            with col6:
                new_notes = st.text_input("用途 / 備註說明")
                
            submit_new = st.form_submit_button("💾 儲存全新藥品入庫", type="primary", use_container_width=True)
            if submit_new and new_name:
                st.success(f"成功新增品項【{new_name} ({new_chi})】共 {new_stock} 顆/件！")

# ---------------------------------------------------------
# Tab 3: 雲端總覽與 Excel 下載
# ---------------------------------------------------------
with tab3:
    st.subheader("📦 目前 Google 雲端藥品總表")
    
    meds_all = load_data("medications")
    
    if not meds_all.empty:
        st.dataframe(meds_all, use_container_width=True)
        
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            meds_all.to_excel(writer, index=False, sheet_name='medications')
        buffer.seek(0)
        
        st.download_button(
            label="🌻 一鍵下載完整藥品庫存 Excel 報表 (.xlsx)",
            data=buffer,
            file_name=f"衛保組藥品庫存總表_{date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    else:
        st.info("資料載入中，或目前尚無庫存資料～")
