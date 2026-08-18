import streamlit as st
import pandas as pd
import io
import os
from datetime import datetime, date

# 嘗試匯入月報表模組 (若 monthly_report.py 尚未建立則給予提示)
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
# 2. 注入 CSS 全域樣式 (選用 Inter 字體、18px+ 大字體、高質感深色卡片)
# -----------------------------------------------------------------------------
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

/* 全域字體設定 (以 Inter 優先，搭配微軟正黑體) */
html, body, [class*="st-"], .stMarkdown, p, span, div {
    font-family: 'Inter', 'Microsoft JhengHei', 'PingFang TC', sans-serif !important;
    font-size: 18px !important;
}

/* 標題字體與大小放大 */
h1 { font-size: 32px !important; font-weight: 700 !important; color: #F8FAFC !important; }
h2 { font-size: 26px !important; font-weight: 700 !important; color: #E2E8F0 !important; }
h3 { font-size: 22px !important; font-weight: 600 !important; color: #CBD5E1 !important; }

/* 表單欄位標籤放大 */
label, .stSelectbox label, .stNumberInput label, .stTextInput label, .stDateInput label, .stRadio label {
    font-size: 18px !important;
    font-weight: 600 !important;
    color: #E2E8F0 !important;
}

/* 輸入框主體文字放大 */
input, select, textarea, .stSelectbox div {
    font-size: 18px !important;
}

/* 一般按鈕樣式 */
.stButton > button {
    font-family: 'Inter', 'Microsoft JhengHei', sans-serif !important;
    font-size: 18px !important;
    font-weight: 600 !important;
    border-radius: 10px !important;
    padding: 10px 20px !important;
}

/* 「雲端報表匯出」頁面專用質感卡片 */
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

/* 下載按鈕超亮眼視覺設計 */
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
# 3. 初始化與讀取資料 (Session State 管理)
# -----------------------------------------------------------------------------
CSV_PATH = "medications_cleaned.csv"
ALT_CSV_PATH = "medications.csv"

def load_initial_data():
    if os.path.exists(CSV_PATH):
        return pd.read_csv(CSV_PATH)
    elif os.path.exists(ALT_CSV_PATH):
        return pd.read_csv(ALT_CSV_PATH)
    else:
        # 預設預備資料 (確保系統無檔案時也能安全啟動)
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
# 4. 側邊欄 (Sidebar) 導覽選單
# -----------------------------------------------------------------------------
with st.sidebar:
    st.title("🏥 衛保組管理系統")
    st.markdown("---")
    
    # 使用側邊欄選項切換頁面
    page = st.radio(
        "📍 請選擇功能頁面",
        ["💊 藥品領用與紀錄", "📦 庫存盤點與校正", "☁️ 雲端報表匯出"],
        index=0
    )
    
    st.markdown("---")
    st.caption("國立臺北大學衛保組 © 115學年度系統")

# -----------------------------------------------------------------------------
# 5. 根據側邊欄選擇渲染不同頁面
# -----------------------------------------------------------------------------

# ==========================================
# 頁面一：💊 藥品領用與紀錄
# ==========================================
if page == "💊 藥品領用與紀錄":
    st.title("💊 藥品領用與登記")
    st.markdown("填寫領藥資訊，系統將自動扣減總庫存並記錄至今日的用量中。")

    df = st.session_state.df
    med_list = df['name'].tolist() if 'name' in df.columns else []

    col1, col2 = st.columns([1, 1])

    with col1:
        selected_med = st.selectbox("選擇藥品", med_list if med_list else ["無藥品資料"])
        qty_used = st.number_input("領取數量", min_value=1, value=1, step=1)
        note = st.text_input("用途 / 備註說明", placeholder="例：發燒、頭痛、去痰")

        if st.button("✅ 確認領藥並儲存", type="primary"):
            if selected_med and selected_med in df['name'].values:
                mask = df['name'] == selected_med
                
                # 扣除總庫存
                df.loc[mask, 'stock'] = df.loc[mask, 'stock'] - qty_used
                
                # 記錄今日領藥數量
                today_key = f"{datetime.now().month}/{datetime.now().day}"
                if today_key not in df.columns:
                    df[today_key] = 0
                df.loc[mask, today_key] = df.loc[mask, today_key] + qty_used
                
                save_data()
                st.success(f"已成功登記領取 {selected_med} 數量：{qty_used}！")
                st.rerun()

    with col2:
        st.subheader("📋 當前藥品庫存總覽")
        display_cols = [c for c in ['name', 'chinese_name', 'stock', 'expiry'] if c in df.columns]
        st.dataframe(df[display_cols], use_container_width=True, height=350)

# ==========================================
# 頁面二：📦 庫存盤點與校正
# ==========================================
elif page == "📦 庫存盤點與校正":
    st.title("📦 庫存盤點與詳細校正")
    st.markdown("針對單一藥品進行詳細庫存資料修訂、盤點數輸入與效期維護。")

    df = st.session_state.df
    med_list = df['name'].tolist() if 'name' in df.columns else []

    if med_list:
        selected_med = st.selectbox("請選擇要校正的藥品", med_list)
        med_row = df[df['name'] == selected_med].iloc[0]

        with st.form("calibration_form"):
            col1, col2 = st.columns(2)
            with col1:
                aug_stock = st.number_input("115年8月剩餘量 (期初庫存)", value=int(med_row.get('aug_stock', med_row.get('stock', 0))))
                current_stock = st.number_input("當前總庫存 (Stock)", value=int(med_row.get('stock', 0)))
                purchased = st.number_input("當月進貨量 (Purchased)", value=int(med_row.get('purchased', 0)))
            with col2:
                expired = st.number_input("過期報銷量 (Expired)", value=int(med_row.get('expired', 0)))
                public_use = st.number_input("公藥使用量 (Public Use)", value=int(med_row.get('public_use', 0)))
                expiry_str = st.text_input("有效期限 (YYYY/MM/DD)", value=str(med_row.get('expiry', '')))

            note_cal = st.text_input("用途 / 備註說明", value="去痰" if selected_med == 'Solmux' else "")

            submitted = st.form_submit_button("🌱 儲存盤點校正紀錄", type="primary")

            if submitted:
                mask = df['name'] == selected_med
                df.loc[mask, 'aug_stock'] = aug_stock
                df.loc[mask, 'stock'] = current_stock
                df.loc[mask, 'purchased'] = purchased
                df.loc[mask, 'expired'] = expired
                df.loc[mask, 'public_use'] = public_use
                df.loc[mask, 'expiry'] = expiry_str
                save_data()
                st.success(f"已成功儲存 {selected_med} 的校正紀錄！")
                st.rerun()

# ==========================================
# 頁面三：☁️ 雲端報表匯出 (僅在此頁面獨家呈現)
# ==========================================
elif page == "☁️ 雲端報表匯出":
    st.title("☁️ 雲端報表匯出中心")

    # 顯示質感 Inter 字體清晰卡片
    st.markdown("""
    <div class="export-card">
        <div class="export-card-title">📊 衛保組 115 學年度藥品使用月報與全學期統計表</div>
        <div class="export-card-desc">
            點擊下方按鈕即可匯出符合校內行政格式之 <b>.xlsx 標準 Excel 報表</b>。<br>
            報表已自動整合每日領藥紀錄、進貨、報銷與全學期動態公式計算。
        </div>
    </div>
    """, unsafe_allow_html=True)

    # 安全讀取資料並產生 Excel 下載按鈕
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
