import streamlit as st
import pandas as pd
import msal
import requests
import os
from urllib.parse import quote
from dotenv import load_dotenv

# 1. โหลดค่าจากไฟล์ .env
load_dotenv()

REQUIRED_ENV_VARS = [
    "TENANT_ID",
    "CLIENT_ID",
    "CLIENT_SECRET",
    "SHAREPOINT_SITE_NAME",
]

# --- ส่วนของการแสดงผล Streamlit (ตั้งค่าหน้าเว็บต้องอยู่บรรทัดแรกสุดของส่วน UI) ---
st.set_page_config(page_title="Vendor Tracking", layout="wide")

# ==========================================
# 🎨 ส่วน CSS สำหรับซ่อนเมนูและ Footer (Clean Mode)
# ==========================================
st.markdown("""
    <style>
        /* 1. ซ่อน Header, Hamburger Menu, Toolbar */
        header[data-testid="stHeader"] {visibility: hidden; display: none !important;}
        .st-emotion-cache-18ni7ap {display: none !important;} /* ซ่อน Toolbar ในบางเวอร์ชัน */
        [data-testid="stToolbar"] {visibility: hidden; display: none !important;}
        #MainMenu {visibility: hidden; display: none !important;}
        
        /* 2. ซ่อน Footer และแถบลิงก์ด้านล่าง */
        footer {visibility: hidden; display: none !important;}
        [data-testid="stFooter"] {visibility: hidden; display: none !important;}
        .stFooter {display: none !important;}
        
        /* 3. ซ่อนปุ่ม Deploy / Manage App / Hosted with Streamlit */
        .stAppDeployButton {display:none !important;}
        [data-testid="stAppDeployButton"] {display:none !important;}
        a[href^="https://streamlit.io/cloud"] {display: none !important;} /* ซ่อนลิงก์ Streamlit Cloud */
        
        /* 4. ซ่อน Decoration และ Status Widget */
        [data-testid="stDecoration"] {display:none !important;}
        [data-testid="stStatusWidget"] {visibility: hidden; display: none !important;}
        
        /* 5. ปรับ Layout ให้ชิดขอบ */
        .block-container {
            padding-top: 1rem !important; 
            padding-bottom: 0rem !important;
        }
        
        /* 6. ซ่อน Element อื่นๆ ที่อาจจะลอยอยู่ (เผื่อไว้) */
        div[class*="viewerBadge"] {display: none !important;}
    </style>
""", unsafe_allow_html=True)

# ==========================================
# ⚙️ ส่วน Backend Functions
# ==========================================

# 2. ฟังก์ชันขอ Token เพื่อเข้าถึง Microsoft Graph
def validate_env():
    missing = [name for name in REQUIRED_ENV_VARS if not os.getenv(name)]
    if missing:
        st.error(
            "ขาดค่าในไฟล์ .env: " + ", ".join(missing)
        )
        return False
    return True


def get_sharepoint_host():
    return os.getenv("SHAREPOINT_HOST", "carchula.sharepoint.com")


def get_sharepoint_file_path():
    file_path = os.getenv("SHAREPOINT_FILE_PATH")
    if file_path:
        return file_path.strip("/")
    folder = os.getenv("SHAREPOINT_FOLDER", "Test Vendor")
    file_name = os.getenv("FILE_NAME", "Payment_Detail_Report.xlsx")
    return f"{folder}/{file_name}".strip("/")


def graph_get_json(url, headers):
    response = requests.get(url, headers=headers)
    if response.status_code >= 400:
        st.error(f"Graph API error {response.status_code}: {response.text}")
        return None
    return response.json()


def get_access_token():
    authority = f"https://login.microsoftonline.com/{os.getenv('TENANT_ID')}"
    app = msal.ConfidentialClientApplication(
        os.getenv('CLIENT_ID'),
        authority=authority,
        client_credential=os.getenv('CLIENT_SECRET')
    )
    # ขอสิทธิ์ระดับ Application (Files.Read.All)
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" in result:
        return result['access_token']
    else:
        st.error("ไม่สามารถเชื่อมต่อ Azure AD ได้ กรุณาเช็ค Credentials")
        return None

# 3. ฟังก์ชันดึงข้อมูลจาก Excel บน SharePoint (อ่าน 2 sheets)
@st.cache_data(ttl=600) # เก็บ Cache ไว้ 10 นาที (ไม่ต้องโหลดใหม่ทุกครั้งที่ User ค้นหา)
def fetch_sharepoint_data():
    if not validate_env():
        return None, None
    token = get_access_token()
    if not token: return None, None
    
    headers = {'Authorization': f'Bearer {token}'}
    
    # 1. ระบุชื่อ Site และ Path ของไฟล์ให้ชัดเจน
    site_name = os.getenv('SHAREPOINT_SITE_NAME') # OARDataGateway2
    file_path = get_sharepoint_file_path()
    sharepoint_host = get_sharepoint_host()
    
    # 2. เปลี่ยน URL เป็นการเจาะจงไฟล์ใน Site
    # เราจะหา Site ID ก่อนแล้วค่อยเจาะไปที่ไฟล์
    site_url = f"https://graph.microsoft.com/v1.0/sites/{sharepoint_host}:/sites/{site_name}"
    site_res = graph_get_json(site_url, headers)
    if not site_res:
        return None, None
    
    if 'id' in site_res:
        site_id = site_res['id']
        # ดึงไฟล์จาก Documents (Shared Documents) > Test Vendor > ไฟล์
        encoded_path = quote(file_path)
        file_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{encoded_path}"
        file_res = graph_get_json(file_url, headers)
        if not file_res:
            return None, None
        
        if '@microsoft.graph.downloadUrl' in file_res:
            download_url = file_res['@microsoft.graph.downloadUrl']
            # กำหนด dtype สำหรับคอลัมน์ที่ต้องเป็น Text เพื่อไม่ให้ pandas แปลงเป็นตัวเลข
            dtype_spec = {
                'เลขที่อ้างอิงรายการ': str,
                'บัญชีหักเงิน': str,
                'บัญชีผู้รับเงิน': str,
                'รหัสธนาคาร': str
            }
            # อ่าน 2 sheets: Payment_Detail_Report และ Show_Column
            df_data = pd.read_excel(download_url, sheet_name='Payment_Detail_Report', dtype=dtype_spec)
            df_show_column = pd.read_excel(download_url, sheet_name='Show_Column')
            return df_data, df_show_column
        else:
            st.error(f"ตรวจพบ Site แต่หาไฟล์ใน Path '{file_path}' ไม่เจอ")
            return None, None
    else:
        st.error(f"เชื่อมต่อ Site '{site_name}' ไม่สำเร็จ ตรวจสอบสิทธิ์ API")
        return None, None

@st.cache_data(ttl=600)
def build_search_index(data, columns):
    if not columns:
        return pd.Series(dtype=str)
    subset = data.loc[:, list(columns)].fillna("")
    return subset.astype(str).agg(" ".join, axis=1).str.lower()


@st.cache_data(ttl=600)
def detect_date_columns(data):
    date_columns = []
    for col in data.columns:
        parsed = pd.to_datetime(data[col], errors="coerce", dayfirst=True)
        if parsed.notna().mean() >= 0.6:
            date_columns.append(col)
    return date_columns


@st.cache_data(ttl=600)
def get_date_series(data, column):
    return pd.to_datetime(data[column], errors="coerce", dayfirst=True)


def format_datetime_columns(df):
    """จัดรูปแบบคอลัมน์วันที่เป็น YYYY-MM-DD เท่านั้น"""
    df_formatted = df.copy()
    for col in df_formatted.columns:
        if pd.api.types.is_datetime64_any_dtype(df_formatted[col]):
            df_formatted[col] = df_formatted[col].dt.strftime('%Y-%m-%d')
    return df_formatted


# --- ส่วนของการแสดงผล Streamlit ---
#st.set_page_config(page_title="Vendor Tracking", layout="wide")
st.title("🔍 ระบบติดตามสถานะการชำระเงิน")

# ปุ่มโหลดข้อมูลใหม่ (ล้าง cache)
col1, col2 = st.columns([0.95, 0.05])
with col2:
    if st.button("🔄", help="โหลดข้อมูลใหม่", key="refresh_btn"):
        st.cache_data.clear()
        st.rerun()

# ดึงข้อมูลมาเก็บไว้ในตัวแปร
if not validate_env():
    st.stop()

df, df_show_column = fetch_sharepoint_data()

if df is not None and df_show_column is not None:
    # กรองคอลัมน์ที่จะแสดงตาม Show_Column sheet
    columns_to_show = df_show_column[df_show_column['Show'].str.upper() == 'YES']['Name'].tolist()
    # เก็บเฉพาะคอลัมน์ที่อยู่ใน df จริงๆ
    columns_to_show = [col for col in columns_to_show if col in df.columns]
    # ปรับ data types ให้ตรงกับ Excel (preserve numeric และ text)
    # pandas จะอ่าน dtype จาก Excel อยู่แล้ว แต่เราจะ handle datetime
    for col in df.columns:
        if df[col].dtype == 'object':
            # ลองแปลงเป็น datetime ถ้าเป็นวันที่
            parsed = pd.to_datetime(df[col], errors='coerce', dayfirst=True)
            if parsed.notna().mean() > 0.6:
                df[col] = parsed
    
    st.sidebar.header("ตัวกรองวันที่")
    
    # กรองวันที่รายการมีผล
    filtered_df = df.copy()
    date_col_1 = "วันที่รายการมีผล"
     
    if date_col_1 in df.columns:
        date_series_1 = get_date_series(df, date_col_1)
        min_date_1 = date_series_1.min()
        max_date_1 = date_series_1.max()
        
        if not pd.isna(min_date_1) and not pd.isna(max_date_1):
            date_range_1 = st.sidebar.date_input(
                f"ช่วง{date_col_1}",
                value=(min_date_1.date(), max_date_1.date()),
                min_value=min_date_1.date(),
                max_value=max_date_1.date(),
                key="date_1"
            )
            if isinstance(date_range_1, tuple) and len(date_range_1) == 2:
                start_date, end_date = date_range_1
                date_mask = date_series_1.dt.date.between(start_date, end_date)
                filtered_df = filtered_df[date_mask]    
 
    search_query = st.text_input(
        "ค้นหา: เลขประจำตัวผู้เสียภาษี",
        placeholder="พิมพ์คำค้นหาที่นี่..."
    )

    result = None
    if search_query:
        query = search_query.strip().lower()
        if not query:
            st.warning("กรุณาพิมพ์คำค้นหาที่ไม่เป็นช่องว่าง")
            st.stop()
        # ค้นหาทุกคอลัมน์
        search_index = build_search_index(filtered_df, tuple(filtered_df.columns))
        mask = search_index.str.contains(query, regex=False, na=False)
        result = filtered_df[mask].copy()
    
    if result is not None and not result.empty:
        # เพิ่ม Running Number
        result.insert(0, 'ลำดับ', range(1, len(result) + 1))
        
        # กรองเฉพาะคอลัมน์ที่ต้องการแสดง
        display_columns = ['ลำดับ'] + [col for col in columns_to_show if col in result.columns]
        result_display = result[display_columns].copy()
        
        st.success(f"พบข้อมูล {len(result_display)} รายการ")
        
        # จัดรูปแบบคอลัมน์วันที่เป็น YYYY-MM-DD
        result_formatted = format_datetime_columns(result_display)
        
        # CSS สำหรับ responsive
        st.markdown("""
        <style>
        @media (max-width: 768px) {
            .stDataFrame { font-size: 12px; }
        }
        </style>
        """, unsafe_allow_html=True)
        
        tab1, tab2 = st.tabs(["📱 รายการ", "📊 ตาราง"])
        
        with tab1:
            # แสดงแบบการ์ด (เป็น default)
            for idx, row in result.iterrows():
                # 1. จัดการยอดเงิน (ใส่ comma และทศนิยม 2 ตำแหน่ง)
                raw_amount = row.get('จำนวนเงิน', 0)
                try:
                    # แปลงเป็น float ก่อน แล้วใช้ format string {:,.2f}
                    amount = f"{float(raw_amount):,.2f}" 
                except (ValueError, TypeError):
                    # กรณีข้อมูลไม่ใช่ตัวเลข (เช่น เป็นขีด '-') ให้แสดงตามเดิม
                    amount = str(raw_amount)

                # 2. จัดการวันที่ (แปลงเป็น YYYY-MM-DD)
                raw_date = row.get('วันที่รายการมีผล', '-')
                if hasattr(raw_date, 'strftime'): 
                    # กรณีเป็น datetime object (จากการใช้ pd.to_datetime)
                    transaction_date = raw_date.strftime('%Y-%m-%d')
                else:
                    # กรณีเป็น String (เช่น '2026-02-10 00:00:00') ให้ตัดเอาแค่ 10 ตัวแรก
                    transaction_date = str(raw_date)[:10]

                # 3. นำตัวแปรที่จัด Format แล้วมาใส่
                recipient_name = row.get('ชื่อผู้รับเงิน', '-')
                card_title = f"🔹 **รายการที่ {int(row.get('ลำดับ', 0))}** - {recipient_name} | **จำนวนเงิน:** {amount} | **วันที่รายการมีผล:** {transaction_date}"

                with st.expander(card_title, expanded=False):
                    cols = st.columns(2)
                    col_idx = 0
                    for col_name in result_formatted.columns:
                        if col_name == 'ลำดับ':
                            continue
                        with cols[col_idx % 2]:
                            value = result_formatted.loc[idx, col_name]
                            if pd.isna(value) or value == "":
                                value = "-"
                            st.markdown(f"**{col_name}:**")
                            st.write(value)
                        col_idx += 1
        
        with tab2:
            # แสดงแบบตาราง
            st.dataframe(
                result_formatted,
                use_container_width=True,
                hide_index=True,
                height=400
            )
    elif result is not None and result.empty:
        st.warning("ไม่พบข้อมูลที่ตรงกับเงื่อนไข")
    else:
        st.info("💡 กรุณากรอกคำค้นหาเพื่อแสดงข้อมูล")