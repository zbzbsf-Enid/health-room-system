import streamlit as st
import pandas as pd
import io
import os
from datetime import datetime

# 嘗試匯入月報表模組
try:
    from monthly_report import generate_monthly_report_excel
except ImportError:
    generate_monthly_report_excel = None

# -----------------------------------------------------------------------------
# 1. 頁面基本配置 (寬頁面模式)
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="衛保組藥品關懷管理系統",
    page_icon="🏥",
    layout="wide"
)

# -----------------------------------------------------------------------------
# 2. 全域 CSS 樣式 (Inter 字體、18px+ 大字體、高質感深色卡片)
# -----------------------------------------------------------------------------
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

html, body, [class*="st-"], .stMarkdown, p, span, div {
    font-family: 'Inter', 'Microsoft JhengHei', 'PingFang TC', sans-serif !important;
    font-size: 18px !important;
}

h1 { font-size: 32px !important; font-weight: 700 !important; color: #F8FAFC !important; }
h2 { font-size: 26px !important; font-weight: 700 !important; color: #E2E8F0 !important; }
h3 { font-size: 22px !important; font-weight: 600 !important; color: #CBD5E1 !important; }

label, .stSelectbox label, .stNumberInput label, .stTextInput label, .stDateInput label, .stRadio label {
    font-size: 18px !important;
    font-weight: 600 !important;
    color: #E2E8F0 !important;
}

input, select, textarea, .stSelectbox div {
    font-size: 18px !important;
}

.stButton > button {
    font-family: 'Inter', 'Microsoft JhengHei', sans-serif !important;
    font-size: 18px !important;
    font-weight: 600 !important;
    border-radius: 10px !important;
    padding: 10px 20px !important;
}

.export-card {
    background: rgba(30, 41, 59, 0.7);
    border: 1.5px solid rgba(255, 255, 255, 0.15);
    border-radius: 16px;
    padding: 28px;
    margin-top: 15px;
    margin-bottom: 25px;
    box-shadow: 0 10px 25px rgba(0, 0, 0, 0.3);
    backdrop-filter: blur(10px);
}

.export-card-title {
    font-family: 'Inter', 'Microsoft JhengHei', sans-serif !important;
    font-size: 26px !important;
    font-weight: 700 !important;
    color: #FFFFFF !important;
    margin-bottom: 14px !important;
    display: flex;
    align-items: center;
    gap: 12px;
}

.export-card-desc {
    font-family: 'Inter', 'Microsoft JhengHei', sans-serif !important;
    font-size: 19px !important;
    line-height: 1.7 !important;
    color: #CBD5E1 !important;
    margin-bottom: 20px !important;
}

div.stDownloadButton > button {
    font-family: 'Inter', 'Microsoft JhengHei', sans-serif !important;
    font-size: 20px !important;
    font-weight: 700 !important;
    padding: 14px 28px !important;
    border-radius: 12px !important;
    background: linear-gradient(135deg, #1F4E78 0%, #2E75B6 100%) !important;
    color: #FFFFFF !important;
    border: none !important;
    box-shadow: 0 6px 16px rgba(31, 78, 120, 0.4) !important;
    transition: all 0.25s ease-in-out !important;
    width: 100% !important;
}

div.stDownloadButton > button:hover {
    background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%) !important;
    box-shadow: 0 8px 22px rgba(37, 99, 235, 0.5) !important;
    transform: translateY(-2px) !important;
}
</style>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# 3. 欄位智慧對齊與相容處理 (解決中文名稱找不到的問題)
# -----------------------------------------------------------------------------
def standardize_dataframe(df):
    """自動辨識各種中文/英文欄位名稱，將其轉為系統標準名稱"""
    col_map = {}
    for col in df.columns:
        c_str = str(col).strip()
        
        # 藥品名稱 (英文/商品名)
        if c_str in ['藥品名稱', '品名', '藥名', '商品名', '英文名稱', 'name', 'med_name']:
            col_map[col] = 'name'
        # 中文名稱
        elif c_str in ['中文名稱', '中文名', '中文藥名', '中文', 'chinese_name', 'cname', 'zh_name']:
            col_map[col] = 'chinese_name'
        # 現有總庫存
        elif c_str in ['當前庫存', '庫存', '現庫存', '總庫存', 'stock', 'current_stock']:
            col_map[col] = 'stock'
        # 期初庫存 / 8月剩餘量
        elif c_str in ['8月剩餘量', '期初庫存', '上月結餘', 'aug_stock', 'august_stock']:
            col_map[col] = 'aug_stock'
        # 有效期限
        elif c_str in ['有效期限', '效期', '到期日', 'expiry', 'exp_date']:
            col_map[col] = 'expiry'
        # 進貨
        elif c_str in ['購入量', '進貨', '進貨量', 'purchased']:
            col_map[col] = 'purchased'
        # 報銷
        elif c_str in ['過期報銷', '報銷', '報廢', 'expired']:
            col_map[col] = 'expired'
        # 公藥
        elif c_str in ['公藥使用', '公藥', 'public_use']:
            col_map[col] = 'public_use'

    df = df.rename(columns=col_map)
    
    # 確保必要欄位存在，若無則自動補齊
    if 'name' not in df.columns:
        df['name'] = "未命名藥品"
    if 'chinese_name' not in df.columns:
        df['chinese_name'] = ""
    else:
        df['chinese_name'] = df['chinese_name'].fillna('')
        
    if 'stock' not in df.columns:
        df['stock'] = 0
    if 'aug_stock' not in df.columns:
        df['aug_stock'] = df['stock']
    if 'expiry' not in df.columns:
        df['expiry'] = ""
    else:
        df['expiry'] = df['expiry'].fillna('')

    for field in ['purchased', 'expired', 'public_use']:
        if field not in df.columns:
            df[field] = 0
        else:
            df[field] = df[field].fillna(0)

    return df

# -----------------------------------------------------------------------------
# 4. 資料載入邏輯
# -----------------------------------------------------------------------------
CSV_PATH = "medications_cleaned.csv"
ALT_CSV_PATH = "medications.csv"

def load_initial_data():
    df = None
    if os.path.exists(CSV_PATH):
        df = pd.read_csv(CSV_PATH)
    elif os.path.exists(ALT_CSV_PATH):
        df = pd.read_csv(ALT_CSV_PATH)
    
    if df is not None and not df.empty:
        return standardize_dataframe(df)
    
    # 預設範例資料
    sept_dates = ['9/1', '9/2', '9/3', '9/4', '9/7', '9/8', '9/9', '9/10', '9/11',
                  '9/14', '9/15', '9/16', '9/17', '9/18', '9/21', '9/22', '9/23',
                  '9/24', '9/25', '9/28', '9/29', '9/30']
    data = {
        'name': ['Panadol', 'Amoxicillin', 'Solmux'],
        'chinese_name': ['普拿疼', '安莫西林', '去痰靈'],
        'aug_stock': [100, 50, 80],
        'stock': [85, 45, 70],
        'purchased': [0, 0, 0],
        'expired': [0, 0, 0],
        'public_use': [0, 0, 0],
        'expiry': ['2028/04/30', '2026/08/18', '2027/12/31']
    }
    for day in sept_dates:
        data[day] = [0, 0, 0]
    return pd.DataFrame(data)

if 'df' not in st.session_state:
    st.session_state.df = load_initial_data()

def save_data():
    st.session_state.df.to_csv(CSV_PATH, index=False, encoding='utf-8-sig')

# -----------------------------------------------------------------------------
# 5. 側邊欄 (Sidebar) 導覽選單與 CSV 匯入區
# -----------------------------------------------------------------------------
with st.sidebar:
    st.title("🏥 衛保組管理系統")
    st.markdown("---")
    
    page = st.radio(
        "📍 請選擇功能頁面",
        ["💊 藥品領用與紀錄", "📦 庫存盤點與校正", "☁️ 雲端報表匯出"],
        index=0
    )
    
    st.markdown("---")
    # ✅ 提供手動上傳/修正舊報表功能
    st.subheader("📤 匯入/修正舊藥品 CSV")
    uploaded_file = st.file_uploader("若中文名稱未顯示，請上傳舊 CSV", type=["csv"])
    if uploaded_file is not None:
        try:
            new_df = pd.read_csv(uploaded_file)
            st.session_state.df = standardize_dataframe(new_df)
            save_data()
            st.success("✅ 舊資料已成功匯入並自動修復中文欄位！")
            st.rerun()
        except Exception as e:
            st.error(f"匯入失敗：{e}")

    st.markdown("---")
    st.caption("國立臺北大學衛保組 © 115學年度系統")

# 產生雙語顯示下拉選單的輔助函式
def get_med_options(df):
    options = []
    for _, row in df.iterrows():
        eng = str(row.get('name', '')).strip()
        chi = str(row.get('chinese_name', '')).strip()
        if chi and chi != 'nan' and chi != 'None':
            options.append(f"{eng} ({chi})")
        else:
            options.append(eng)
    return options

# -----------------------------------------------------------------------------
# 6. 頁面渲染邏輯
# -----------------------------------------------------------------------------
df = st.session_state.df

# ==========================================
# 頁面一：💊 藥品領用與紀錄
# ==========================================
if page == "💊 藥品領用與紀錄":
    st.title("💊 藥品領用與登記")
    st.markdown("填寫領藥資訊，系統將自動扣減總庫存並記錄至今日用量。")

    col1, col2 = st.columns([1, 1])

    options = get_med_options(df)

    with col1:
        selected_option = st.selectbox("選擇藥品", options if options else ["無藥品資料"])
        qty_used = st.number_input("領取數量", min_value=1, value=1, step=1)
        note = st.text_input("用途 / 備註說明", placeholder="例：發燒、頭痛、去痰")

        if st.button("✅ 確認領藥並儲存", type="primary"):
            if selected_option and options:
                idx = options.index(selected_option)
                med_name = df.iloc[idx]['name']
                mask = df['name'] == med_name
                
                df.loc[mask, 'stock'] = df.loc[mask, 'stock'] - qty_used
                
                today_key = f"{datetime.now().month}/{datetime.now().day}"
                if today_key not in df.columns:
                    df[today_key] = 0
                df.loc[mask, today_key] = df.loc[mask, today_key] + qty_used
                
                save_data()
                st.success(f"已成功登記領取 {selected_option} 數量：{qty_used}！")
                st.rerun()

    with col2:
        st.subheader("📋 當前藥品庫存總覽")
        # 美化顯示表格，明確標示中文欄位
        show_df = pd.DataFrame()
        show_df["藥品名稱 (英文)"] = df['name']
        show_df["中文名稱"] = df['chinese_name']
        show_df["目前庫存"] = df['stock']
        show_df["有效期限"] = df['expiry']
        st.dataframe(show_df, use_container_width=True, height=380)

# ==========================================
# 頁面二：📦 庫存盤點與校正
# ==========================================
elif page == "📦 庫存盤點與校正":
    st.title("📦 庫存盤點與詳細校正")
    st.markdown("針對單一藥品進行詳細庫存資料修訂、盤點數輸入與效期維護。")

    options = get_med_options(df)

    if options:
        selected_option = st.selectbox("請選擇要校正的藥品", options)
        idx = options.index(selected_option)
        med_row = df.iloc[idx]
        selected_med = med_row['name']

        with st.form("calibration_form"):
            col1, col2 = st.columns(2)
            with col1:
                eng_name = st.text_input("藥品名稱 (商品名/英文)", value=str(med_row.get('name', '')))
                chi_name = st.text_input("中文名稱", value=str(med_row.get('chinese_name', '')))
                aug_stock = st.number_input("115年8月剩餘量 (期初庫存)", value=int(med_row.get('aug_stock', med_row.get('stock', 0))))
                current_stock = st.number_input("當前總庫存 (Stock)", value=int(med_row.get('stock', 0)))
            with col2:
                purchased = st.number_input("當月進貨量 (Purchased)", value=int(med_row.get('purchased', 0)))
                expired = st.number_input("過期報銷量 (Expired)", value=int(med_row.get('expired', 0)))
                public_use = st.number_input("公藥使用量 (Public Use)", value=int(med_row.get('public_use', 0)))
                expiry_str = st.text_input("有效期限 (YYYY/MM/DD)", value=str(med_row.get('expiry', '')))

            submitted = st.form_submit_button("🌱 儲存盤點校正紀錄", type="primary")

            if submitted:
                mask = df['name'] == selected_med
                df.loc[mask, 'name'] = eng_name
                df.loc[mask, 'chinese_name'] = chi_name
                df.loc[mask, 'aug_stock'] = aug_stock
                df.loc[mask, 'stock'] = current_stock
                df.loc[mask, 'purchased'] = purchased
                df.loc[mask, 'expired'] = expired
                df.loc[mask, 'public_use'] = public_use
                df.loc[mask, 'expiry'] = expiry_str
                save_data()
                st.success(f"已成功儲存 {eng_name} ({chi_name}) 的校正紀錄！")
                st.rerun()

# ==========================================
# 頁面三：☁️ 雲端報表匯出
# ==========================================
elif page == "☁️ 雲端報表匯出":
    st.title("☁️ 雲端報表匯出中心")

    st.markdown("""
    <div class="export-card">
        <div class="export-card-title">📊 衛保組 115 學年度藥品使用月報與全學期統計表</div>
        <div class="export-card-desc">
            點擊下方按鈕即可匯出符合校內行政格式之 <b>.xlsx 標準 Excel 報表</b>。<br>
            報表已自動整合每日領藥紀錄、進貨、報銷與全學期動態公式計算。
        </div>
    </div>
    """, unsafe_allow_html=True)

    if generate_monthly_report_excel is not None:
        try:
            excel_bytes = generate_monthly_report_excel(st.session_state.df)
            st.download_button(
                label="📥 點此下載 115 學年度用藥月報與全學期統計表 (.xlsx)",
                data=excel_bytes,
                file_name="115學年度上學期用藥月報與學期統計表.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"⚠️ 產生報表時發生錯誤：{e}")
    else:
        st.warning("⚠️ 找不到 `monthly_report.py` 模組，請確定該檔案已建立並位於專案總資料夾中。")
