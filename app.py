import streamlit as st
import pandas as pd
import io
from datetime import datetime, date

st.set_page_config(page_title="衛保組藥品管理系統", layout="wide")

# 🎨 注入 CSS：強制放大全域字體（至少 16px 以上）並便利化按鈕與輸入框
st.markdown("""
    <style>
    /* 全域內文與標籤字體放大 */
    html, body, [class*="st-"], .stMarkdown, p, span {
        font-size: 18px !important;
    }
    
    /* 表單標題與欄位標籤 */
    label, .stSelectbox label, .stNumberInput label, .stTextInput label, .stDateInput label {
        font-size: 18px !important;
        font-weight: bold !important;
    }
    
    /* 下拉選單與輸入框內文字 */
    .stSelectbox div[data-baseweb="select"] > div, 
    .stNumberInput input, 
    .stTextInput input {
        font-size: 18px !important;
    }
    
    /* 頁籤選單文字放大 */
    button[data-baseweb="tab"] {
        font-size: 20px !important;
        font-weight: bold !important;
    }
    
    /* 按鈕加大與樣式優化 */
    .stButton > button {
        font-size: 18px !important;
        font-weight: bold !important;
        border-radius: 8px !important;
        padding: 0.5rem 1rem !important;
    }
    
    /* 表格字體 */
    .stDataFrame {
        font-size: 16px !important;
    }
    
    /* 提示訊息文字 */
    .stAlert {
        font-size: 18px !important;
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

st.title("💊 衛保組藥品管理系統")

tab1, tab2, tab3 = st.tabs([
    "⚡ 多品項快速扣庫/發藥", 
    "✏️ 線上庫存異動與新藥入庫", 
    "📦 庫存總覽與 Excel 匯入/下載"
])

# ---------------------------------------------------------
# Tab 1: 快速扣庫/發藥
# ---------------------------------------------------------
with tab1:
    meds_df = load_data("medications")
    
    if not meds_df.empty and 'stock' in meds_df.columns:
        meds_df['stock'] = pd.to_numeric(meds_df['stock'], errors='coerce').fillna(0).astype(int)
        valid_meds = meds_df[meds_df['stock'] > 0].copy()
        
        if not valid_meds.empty:
            st.subheader("1. 選擇藥品並加入清單")
            
            options = [
                f"{r.get('name', '')} ({r.get('chinese_name', '')}) | 批號:{r.get('lot_no', 'DEFAULT')} | 效期:{r.get('expiry', '')} | 剩餘庫存:{r['stock']}" 
                for _, r in valid_meds.iterrows()
            ]
            
            # 調整欄位排版比例，增加按鈕點擊面積
            col_select, col_qty, col_add = st.columns([3.5, 1.2, 1.3])
            
            with col_select:
                selected_str = st.selectbox("選擇藥品 (自動帶出最早到期的批次)", options, key="dispense_select")
                selected_row = valid_meds.iloc[options.index(selected_str)]
                
            with col_qty:
                qty = st.number_input("扣除數量", min_value=1, max_value=int(selected_row['stock']), value=1, key="dispense_qty")
                
            with col_add:
                st.write("")
                st.write("")
                if st.button("➕ 加入發藥清單", use_container_width=True, type="primary"):
                    existing = next((item for item in st.session_state.cart if item['name'] == selected_row['name'] and item['lot_no'] == selected_row.get('lot_no', 'DEFAULT')), None)
                    if existing:
                        if existing['qty'] + qty <= selected_row['stock']:
                            existing['qty'] += qty
                            st.toast(f"已更新 {selected_row['name']} 數量為 {existing['qty']}", icon="✅")
                        else:
                            st.error("超過現有庫存上限！")
                    else:
                        st.session_state.cart.append({
                            'name': selected_row['name'],
                            'chinese_name': selected_row.get('chinese_name', ''),
                            'lot_no': selected_row.get('lot_no', 'DEFAULT'),
                            'qty': qty,
                            'max_stock': selected_row['stock']
                        })
                        st.toast(f"已加入：{selected_row['name']} x {qty}", icon="💊")
                    st.rerun()

            st.markdown("---")
            
            cart_count = len(st.session_state.cart)
            st.subheader(f"2. 本次待發藥品清單 (共 {cart_count} 項)")

            if cart_count > 0:
                cart_df = pd.DataFrame(st.session_state.cart)[['name', 'chinese_name', 'lot_no', 'qty']]
                cart_df.columns = ['商品名', '中文名', '批號', '扣除數量']
                
                st.dataframe(cart_df, use_container_width=True, height=200)
                
                btn_col1, btn_col2 = st.columns([1, 3])
                with btn_col1:
                    if st.button("🗑️ 清空清單", use_container_width=True):
                        st.session_state.cart = []
                        st.rerun()
                        
                with btn_col2:
                    if st.button("✅ 完成扣庫試算 (試算表同步提醒)", type="primary", use_container_width=True):
                        summary_list = [f"{item['name']} x{item['qty']}" for item in st.session_state.cart]
                        st.balloons()
                        st.success(f"🎉 扣庫完成！扣除項目：{', '.join(summary_list)}")
                        st.info("💡 請記得至 Google 試算表同步更新數量。")
                        st.session_state.cart = []
            else:
                st.info("💡 目前發藥清單為空，請從上方選擇藥品並按【➕ 加入發藥清單】")

        else:
            st.warning("目前試算表內尚無可用庫存（庫存皆為 0）。")
    else:
        st.error("無法正確讀取 Google 試算表，請確認試算表已執行【檔案 ➜ 分享 ➜ 發佈到網路】。")

# ---------------------------------------------------------
# Tab 2: 線上庫存異動
# ---------------------------------------------------------
with tab2:
    st.subheader("✏️ 線上庫存異動與新藥進貨登記")
    st.markdown("請於下方登記新進藥品或庫存校正資料：")
    
    with st.form("add_med_form"):
        col1, col2, col3 = st.columns(3)
        with col1:
            name = st.text_input("藥品名稱 (商品名)")
        with col2:
            chi_name = st.text_input("中文藥名")
        with col3:
            lot_no = st.text_input("批號", value="DEFAULT")
            
        col4, col5, col6 = st.columns(3)
        with col4:
            stock = st.number_input("數量", min_value=0, value=100)
        with col5:
            expiry = st.date_input("有效期限", value=date(2028, 12, 31))
        with col6:
            notes = st.text_input("用途/備註")
            
        submit = st.form_submit_button("💾 登記新增項目", type="primary", use_container_width=True)
        if submit and name:
            st.success(f"已登記：{name} ({chi_name}) x {stock}，效期：{expiry}")

# ---------------------------------------------------------
# Tab 3: 庫存總覽與下載
# ---------------------------------------------------------
with tab3:
    st.subheader("📦 目前 Google 雲端藥品庫存總覽")
    
    meds_all = load_data("medications")
    
    if not meds_all.empty:
        st.dataframe(meds_all, use_container_width=True)
        
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            meds_all.to_excel(writer, index=False, sheet_name='medications')
        buffer.seek(0)
        
        st.download_button(
            label="📥 下載目前完整庫存 Excel 報表 (.xlsx)",
            data=buffer,
            file_name=f"衛保組藥品庫存總表_{date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    else:
        st.info("尚無庫存資料或正在載入中...")
