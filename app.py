import streamlit as st
import pandas as pd
import io
import os
import re
from datetime import datetime

# 嘗試匯入月報表模組
try:
    from monthly_report import generate_monthly_report_excel
except ImportError:
    generate_monthly_report_excel = None

# -----------------------------------------------------------------------------
# 1. 頁面基本配置
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="衛保組藥品關懷管理系統",
    page_icon="🏥",
    layout="wide"
)

# -----------------------------------------------------------------------------
# 2. 全域 CSS 樣式 (適中 15px/16px 字體、質感深色風格)
# -----------------------------------------------------------------------------
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

html, body, [class*="st-"], .stMarkdown, p, span, div {
    font-family: 'Inter', 'Microsoft JhengHei', 'PingFang TC', sans-serif !important;
    font-size: 16px !important;
}

h1 { font-size: 28px !important; font-weight: 700 !important; color: #F8FAFC !important; }
h2 { font-size: 22px !important; font-weight: 700 !important; color: #E2E8F0 !important; }
h3 { font-size: 18px !important; font-weight: 600 !important; color: #CBD5E1 !important; }

label, .stSelectbox label, .stMultiSelect label, .stNumberInput label, .stTextInput label {
    font-size: 16px !important;
    font-weight: 600 !important;
    color: #E2E8F0 !important;
}

/* 🎯 側邊欄與選單專用：字體縮小至 15px */
section[data-testid="stSidebar"] *,
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] p,
section[data-testid="stSidebar"] span {
    font-size: 15px !important;
}

/* 🎯 下拉與多選選單內容文字 15px */
input, select, textarea, .stSelectbox div, div[data-baseweb="select"] * {
    font-size: 15px !important;
}

/* 按鈕樣式 */
.stButton > button {
    font-family: 'Inter', 'Microsoft JhengHei', sans-serif !important;
    font-size: 16px !important;
    font-weight: 600 !important;
    border-radius: 10px !important;
    padding: 8px 18px !important;
}

/* 批次領藥卡片外框 */
.batch-dispense-box {
    background: rgba(30, 41, 59, 0.5);
    border: 1px solid rgba(255, 255, 255, 0.1);
    border-radius: 12px;
    padding: 16px;
    margin-bottom: 15px;
}

/* 雲端報表匯出卡片 */
.export-card {
    background: rgba(30, 41, 59, 0.7);
    border: 1.5px solid rgba(255, 255, 255, 0.15);
    border-radius: 16px;
    padding: 24px;
    margin-top: 15px;
    margin-bottom: 25px;
    box-shadow: 0 10px 25px rgba(0, 0, 0, 0.3);
    backdrop-filter: blur(10px);
}

.export-card-title {
    font-family: 'Inter', 'Microsoft JhengHei', sans-serif !important;
    font-size: 22px !important;
    font-weight: 700 !important;
    color: #FFFFFF !important;
    margin-bottom: 12px !important;
}

.export-card-desc {
    font-family: 'Inter', 'Microsoft JhengHei', sans-serif !important;
    font-size: 16px !important;
    line-height: 1.6 !important;
    color: #CBD5E1 !important;
    margin-bottom: 15px !important;
}

div.stDownloadButton > button {
    font-family: 'Inter', 'Microsoft JhengHei', sans-serif !important;
    font-size: 18px !important;
    font-weight: 700 !important;
    padding: 12px 24px !important;
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
# 3. 欄位智慧解析與正規化
# -----------------------------------------------------------------------------
def standardize_dataframe(df):
    """解析並正規化來自 Google 試算表或本地 CSV 的資料"""
    col_map = {}
    for col in df.columns:
        c_str = str(col).strip()
        if c_str in ['name', '藥品名稱', '品名', '藥名']:
            col_map[col] = 'name'
        elif c_str.startswith('chinese') or c_str in ['中文名稱', '中文名', '中文']:
            col_map[col] = 'chinese_name'
        elif c_str in ['stock', '當前庫存', '庫存', '現庫存', '總庫存']:
            col_map[col] = 'stock'
        elif c_str in ['aug_stock', '8月剩餘量', '期初庫存']:
            col_map[col] = 'aug_stock'
        elif c_str in ['expiry', '有效期限', '效期']:
            col_map[col] = 'expiry'
        elif c_str in ['notes', '用途', '備註']:
            col_map[col] = 'notes'
        elif c_str in ['purchased', '購入量', '進貨']:
            col_map[col] = 'purchased'
        elif c_str in ['expired', '過期報銷', '報銷']:
            col_map[col] = 'expired'
        elif c_str in ['public_use', '公藥使用', '公藥']:
            col_map[col] = 'public_use'

    df = df.rename(columns=col_map)
    
    if 'name' not in df.columns: df['name'] = "未命名藥品"
    if 'chinese_name' not in df.columns: df['chinese_name'] = ""
    else: df['chinese_name'] = df['chinese_name'].fillna('')
    
    if 'stock' not in df.columns: df['stock'] = 0
    else: df['stock'] = pd.to_numeric(df['stock'], errors='coerce').fillna(0).astype(int)
    
    if 'aug_stock' not in df.columns: df['aug_stock'] = df['stock']
    if 'expiry' not in df.columns: df['expiry'] = ""
    else: df['expiry'] = df['expiry'].fillna('').astype(str)
    
    if 'notes' not in df.columns: df['notes'] = ""
    else: df['notes'] = df['notes'].fillna('')

    for col_name in ['purchased', 'expired', 'public_use']:
        if col_name not in df.columns:
            df[col_name] = 0
        else:
            df[col_name] = pd.to_numeric(df[col_name], errors='coerce').fillna(0).astype(int)

    return df

def convert_gsheet_url_to_csv(url):
    """將 Google 試算表公開網址轉為可下載 CSV 的網址"""
    pattern = r"/d/([a-zA-Z0-9-_]+)"
    match = re.search(pattern, url)
    if match:
        spreadsheet_id = match.group(1)
        return f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=csv"
    return url

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
    
    # 預設備用資料
    data = {
        'name': ['Actein 600mg', 'Amoxicillin 500mg', 'Ancogen'],
        'chinese_name': ['愛克痰發泡錠', '安莫西林', '安可腱'],
        'aug_stock': [643, 1403, 588],
        'stock': [643, 1403, 588],
        'purchased': [0, 0, 0],
        'expired': [0, 0, 0],
        'public_use': [0, 0, 0],
        'expiry': ['2028-04-30', '2026-08-31', '2027-03-31'],
        'notes': ['去痰', '抗生素', '骨骼肌鬆弛']
    }
    return pd.DataFrame(data)

if 'df' not in st.session_state:
    st.session_state.df = load_initial_data()

def save_data():
    st.session_state.df.to_csv(CSV_PATH, index=False, encoding='utf-8-sig')

# -----------------------------------------------------------------------------
# 5. 側邊欄 (Sidebar) 選單區
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
    st.subheader("🔗 串接 Google 雲端試算表")
    gsheet_url_input = st.text_input("貼上 Google 試算表連結：", placeholder="https://docs.google.com/spreadsheets/d/...")
    
    if st.button("🔄 同步雲端試算表資料"):
        if gsheet_url_input.strip():
            try:
                csv_export_url = convert_gsheet_url_to_csv(gsheet_url_input.strip())
                new_df = pd.read_csv(csv_export_url)
                st.session_state.df = standardize_dataframe(new_df)
                save_data()
                st.success("✅ 已成功載入 Google 試算表最新藥品資料！")
                st.rerun()
            except Exception as e:
                st.error(f"無法同步試算表，請確認連結已設定為『知道連結的人皆可檢視』！\n錯誤：{e}")
        else:
            st.warning("請先輸入網址！")

    st.markdown("---")
    st.subheader("📤 上傳本地 CSV / Excel 檔")
    uploaded_file = st.file_uploader("手動上傳藥品清單", type=["csv", "xlsx"])
    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith('.csv'):
                new_df = pd.read_csv(uploaded_file)
            else:
                new_df = pd.read_excel(uploaded_file)
            st.session_state.df = standardize_dataframe(new_df)
            save_data()
            st.success("✅ 檔案上傳成功並已匯入！")
            st.rerun()
        except Exception as e:
            st.error(f"匯入失敗：{e}")

    st.markdown("---")
    st.caption("國立臺北大學衛保組 © 115學年度系統")

# 格式化選單選項名稱 (傳回列索引, 避免藥名重複時混淆)
def format_med_option(idx, df_data):
    row = df_data.iloc[idx]
    eng = str(row.get('name', '')).strip()
    chi = str(row.get('chinese_name', '')).strip()
    exp = str(row.get('expiry', '')).strip()
    stk = str(row.get('stock', 0)).strip()
    
    res = eng
    if chi and chi not in ['nan', 'None', '']:
        res += f" ({chi})"
    if exp and exp not in ['nan', 'None', '']:
        res += f" | 效期:{exp}"
    res += f" (庫存:{stk})"
    return res

# -----------------------------------------------------------------------------
# 6. 頁面渲染邏輯
# -----------------------------------------------------------------------------
df = st.session_state.df

# ==========================================
# 頁面一：💊 藥品領用與紀錄 (升級：多選批次領藥模式)
# ==========================================
if page == "💊 藥品領用與紀錄":
    st.title("💊 藥品領用與登記")
    st.markdown("點選下方搜尋欄可**選擇一種或多種藥品**，設定數量後即可一次完成登記與庫存扣減。")

    col1, col2 = st.columns([1.1, 1])

    with col1:
        options_indices = list(range(len(df)))
        
        # 多選下拉選單
        selected_indices = st.multiselect(
            "選擇本次領取的所有藥品 (可同時選擇多項)",
            options=options_indices,
            format_func=lambda idx: format_med_option(idx, df),
            placeholder="請點擊或輸入藥名/中文名稱進行搜尋..."
        )

        items_to_dispense = []

        if selected_indices:
            st.markdown("---")
            st.markdown("##### 📝 設定各領取藥品數量")
            
            # 為每一個選中的藥品產生動態數量調整框
            for i, idx in enumerate(selected_indices):
                row_data = df.iloc[idx]
                med_name = row_data['name']
                chi_name = row_data['chinese_name']
                curr_stock = int(row_data['stock'])
                
                c_info, c_qty = st.columns([2.5, 1.2])
                with c_info:
                    st.write(f"**{med_name}**")
                    st.caption(f"{chi_name}｜當前庫存：{curr_stock}")
                with c_qty:
                    qty = st.number_input(
                        "數量",
                        min_value=1,
                        max_value=max(1, curr_stock),
                        value=1,
                        step=1,
                        key=f"qty_input_{idx}_{i}"
                    )
                items_to_dispense.append((idx, med_name, chi_name, qty))

            st.markdown("---")
            note = st.text_input("用途 / 備註說明 (此批領藥共通)", placeholder="例：發燒感冒、頭痛、去痰、傷口處置")

            if st.button("✅ 一鍵批次登記領取", type="primary"):
                today_key = f"{datetime.now().month}/{datetime.now().day}"
                if today_key not in df.columns:
                    df[today_key] = 0
                
                summary_logs = []
                for idx, med_name, chi_name, qty in items_to_dispense:
                    mask = df.index == idx
                    df.loc[mask, 'stock'] = df.loc[mask, 'stock'] - qty
                    df.loc[mask, today_key] = df.loc[mask, today_key] + qty
                    summary_logs.append(f"{med_name} ({chi_name}) x{qty}")
                
                save_data()
                st.success(f"🎉 已成功完成批次登記領藥！\n領取項目：{', '.join(summary_logs)}")
                st.rerun()
        else:
            st.info("💡 請先在上方的選單中點選或搜尋要領取的藥品。")

    with col2:
        st.subheader("📋 當前藥品庫存總覽")
        show_df = pd.DataFrame()
        show_df["藥品名稱 (英文)"] = df['name']
        show_df["中文名稱"] = df['chinese_name']
        show_df["目前庫存"] = df['stock']
        show_df["有效期限"] = df['expiry']
        show_df["用途/備註"] = df['notes']
        st.dataframe(show_df, use_container_width=True, height=480)

# ==========================================
# 頁面二：📦 庫存盤點與校正
# ==========================================
elif page == "📦 庫存盤點與校正":
    st.title("📦 庫存盤點與詳細校正")
    st.markdown("針對單一藥品進行詳細庫存資料修訂、盤點數輸入與效期維護。")

    options_indices = list(range(len(df)))

    if options_indices:
        selected_idx = st.selectbox(
            "請選擇要校正的藥品",
            options=options_indices,
            format_func=lambda idx: format_med_option(idx, df)
        )
        med_row = df.iloc[selected_idx]

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
                expiry_str = st.text_input("有效期限 (YYYY-MM-DD)", value=str(med_row.get('expiry', '')))

            note_cal = st.text_input("用途 / 備註說明", value=str(med_row.get('notes', '')))

            submitted = st.form_submit_button("🌱 儲存盤點校正紀錄", type="primary")

            if submitted:
                mask = df.index == selected_idx
                df.loc[mask, 'name'] = eng_name
                df.loc[mask, 'chinese_name'] = chi_name
                df.loc[mask, 'aug_stock'] = aug_stock
                df.loc[mask, 'stock'] = current_stock
                df.loc[mask, 'purchased'] = purchased
                df.loc[mask, 'expired'] = expired
                df.loc[mask, 'public_use'] = public_use
                df.loc[mask, 'expiry'] = expiry_str
                df.loc[mask, 'notes'] = note_cal
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
                label="📥 點此下載 115 學年度用藥月報與學期統計表 (.xlsx)",
                data=excel_bytes,
                file_name="115學年度上學期用藥月報與學期統計表.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"⚠️ 產生報表時發生錯誤：{e}")
    else:
        st.warning("⚠️ 找不到 `monthly_report.py` 模組，請確定該檔案已建立並位於專案總資料夾中。")
