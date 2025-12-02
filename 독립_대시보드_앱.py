"""
잔여객실 현황 독립 대시보드
- 개발자 없이 실행 가능
- DB 직접 연결
- 웹 브라우저에서 실행
"""

import streamlit as st
import pyodbc
import pandas as pd
from datetime import datetime, timedelta

# ⚠️ set_page_config()는 반드시 첫 번째 Streamlit 명령이어야 함!
st.set_page_config(page_title="객실 현황 대시보드", layout="wide", initial_sidebar_state="collapsed")

# ============================================================
# 📝 DB 설정 - Streamlit Secrets 사용
# ============================================================
try:
    # Streamlit Cloud 또는 로컬 secrets.toml 사용
    DB_CONFIG = {
        'server': st.secrets["database"]["server"],
        'base_database': st.secrets["database"]["base_database"],
        'cruise_database': st.secrets["database"]["cruise_database"],
        'username': st.secrets["database"]["username"],
        'password': st.secrets["database"]["password"],
    }
except Exception as e:
    # Secrets 없을 때 기본값 (개발용)
    st.error("⚠️ DB 설정을 찾을 수 없습니다. .streamlit/secrets.toml 파일을 확인하세요.")
    DB_CONFIG = {
        'server': '',
        'base_database': '',
        'cruise_database': '',
        'username': '',
        'password': '',
    }
# ============================================================

# Session state 초기화 (필요시)
st.markdown("""
<div style="text-align: center; padding: 60px 0 40px 0;">
    <h1 style="font-size: 48px; font-weight: 300; letter-spacing: -1px; color: #0a0a0a; margin-bottom: 12px;">
        NEOHELIOS CRUISE
    </h1>
    <p style="font-size: 16px; font-weight: 400; color: #6b6b6b; letter-spacing: 0.5px; text-transform: uppercase;">
        Cabin Availability Dashboard
    </p>
</div>
""", unsafe_allow_html=True)

# Palantir 스타일 정의 (클린 화이트 + 반응형)
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&display=swap');
    
    /* 전체 배경 - Palantir 스타일 */
    .stApp {
        background: #fafafa;
        color: #0a0a0a;
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
    }
    
    /* 메인 컨테이너 */
    .main {
        background: #fafafa;
        padding: 2rem;
        max-width: 1600px;
        margin: 0 auto;
    }
    
    /* 제목 스타일 */
    h1 {
        color: #1a1a1a;
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
        font-weight: 600;
        letter-spacing: -0.5px;
        margin-bottom: 10px;
    }
    
    h3 {
        color: #2d2d2d;
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
        font-weight: 500;
        font-size: 14px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        margin-bottom: 16px;
    }
    
    /* 버튼 스타일 - Palantir 느낌 */
    .stButton > button {
        background: #0a0a0a;
        color: #ffffff;
        font-weight: 500;
        border: none;
        border-radius: 2px;
        padding: 14px 28px;
        font-size: 15px;
        letter-spacing: 0.5px;
        text-transform: uppercase;
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.12);
    }
    
    .stButton > button:hover {
        background: #2a2a2a;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
        transform: translateY(-1px);
    }
    
    /* 다운로드 버튼 - 흰색 배경에 보더 */
    .stDownloadButton > button {
        background: #ffffff;
        color: #0a0a0a;
        border: 1px solid #d0d0d0;
        font-weight: 500;
        border-radius: 2px;
        padding: 18px 36px;
        font-size: 16px;
        letter-spacing: 0.5px;
        text-transform: uppercase;
        width: 100%;
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
    }
    
    .stDownloadButton > button:hover {
        background: #0a0a0a;
        color: #ffffff;
        border-color: #0a0a0a;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
        transform: translateY(-1px);
    }
    
    /* 입력 필드 - 미니멀 스타일 */
    .stSelectbox > div > div,
    .stDateInput > div > div,
    .stTextInput > div > div {
        border-radius: 2px;
        border: 1px solid #d0d0d0;
        background: #ffffff;
        transition: all 0.2s ease;
    }
    
    .stSelectbox > div > div:hover,
    .stDateInput > div > div:hover,
    .stTextInput > div > div:hover {
        border-color: #0a0a0a;
    }
    
    /* 라벨 스타일 */
    .stSelectbox label,
    .stDateInput label,
    .stTextInput label {
        font-size: 14px;
        font-weight: 600;
        letter-spacing: 0.5px;
        text-transform: uppercase;
        color: #0a0a0a;
    }
    
    /* 입력 필드 내 텍스트 */
    .stSelectbox input,
    .stSelectbox select,
    .stSelectbox div[data-baseweb="select"] > div,
    .stSelectbox div[data-baseweb="select"] span,
    .stDateInput input,
    .stTextInput input {
        color: #0a0a0a !important;
        font-size: 17px !important;
        background-color: #ffffff !important;
    }
    
    /* Disabled 필드 (선박) */
    .stTextInput input:disabled {
        color: #0a0a0a !important;
        background-color: #f5f5f5 !important;
        -webkit-text-fill-color: #0a0a0a !important;
        opacity: 1 !important;
    }
    
    /* 드롭다운 메뉴 */
    [role="listbox"],
    [data-baseweb="menu"],
    [data-baseweb="popover"] > div,
    .stSelectbox [role="option"] {
        background-color: #ffffff !important;
    }
    
    /* 드롭다운 옵션 텍스트 */
    [role="option"],
    [data-baseweb="menu"] li,
    .stSelectbox li {
        background-color: #ffffff !important;
        color: #0a0a0a !important;
        font-size: 16px !important;
    }
    
    /* 드롭다운 옵션 호버 */
    [role="option"]:hover,
    [data-baseweb="menu"] li:hover {
        background-color: #f0f0f0 !important;
        color: #0a0a0a !important;
    }
    
    /* Success/Info 메시지 */
    .stSuccess {
        background: #f0f8f0;
        border-left: 4px solid #2d7a2d;
        padding: 12px 16px;
        border-radius: 2px;
    }
    
    .stInfo {
        background: #f0f4f8;
        border-left: 4px solid #2d5a7a;
        padding: 12px 16px;
        border-radius: 2px;
    }
    
    /* 탭 스타일 - 밑줄 스타일 + 가독성 최우선 */
    .stTabs [data-baseweb="tab-list"] {
        gap: 24px;
        background-color: transparent;
        border-bottom: 2px solid #e0e0e0;
        padding-bottom: 0;
        margin-bottom: 30px;
    }
    
    /* 비활성 탭 - 검은색 텍스트 */
    .stTabs [data-baseweb="tab"] {
        background-color: transparent !important;
        border: none !important;
        border-bottom: 4px solid transparent !important;
        border-radius: 0 !important;
        padding: 20px 32px !important;
        transition: all 0.2s ease;
    }
    
    .stTabs button[data-baseweb="tab"] {
        color: #6b6b6b !important;
        font-size: 36px !important;
        font-weight: 600 !important;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    
    .stTabs [data-baseweb="tab"]:hover button {
        color: #0a0a0a !important;
    }
    
    /* 활성 탭 - 검은색 텍스트 + 밑줄 */
    .stTabs [aria-selected="true"] {
        background-color: transparent !important;
        border-bottom: 4px solid #0a0a0a !important;
    }
    
    .stTabs [aria-selected="true"] button {
        color: #0a0a0a !important;
        font-weight: 700 !important;
    }
    
    .stTabs [data-baseweb="tab-panel"] {
        padding-top: 30px;
    }
    
    /* 반응형 테이블 컨테이너 */
    .responsive-table-container {
        overflow-x: auto;
        -webkit-overflow-scrolling: touch;
        margin: 20px 0;
    }
    
    .responsive-table-container table {
        min-width: 100%;
    }
    
    /* 클릭 가능한 셀 */
    .clickable-cell {
        cursor: pointer;
        transition: all 0.2s ease;
    }
    
    .clickable-cell:hover {
        opacity: 0.8;
        transform: scale(1.05);
    }
    
    /* 모달 스타일 */
    .modal {
        display: none;
        position: fixed;
        z-index: 9999;
        left: 0;
        top: 0;
        width: 100%;
        height: 100%;
        overflow: auto;
        background-color: rgba(0, 0, 0, 0.6);
        animation: fadeIn 0.3s ease;
    }
    
    .modal-content {
        background-color: #ffffff;
        margin: 5% auto;
        padding: 0;
        border-radius: 8px;
        width: 90%;
        max-width: 900px;
        box-shadow: 0 8px 32px rgba(0, 0, 0, 0.3);
        animation: slideDown 0.3s ease;
    }
    
    .modal-header {
        background: #1a1a1a;
        color: white;
        padding: 20px 24px;
        border-radius: 8px 8px 0 0;
        display: flex;
        justify-content: space-between;
        align-items: center;
    }
    
    .modal-header h2 {
        margin: 0;
        font-size: 20px;
        font-weight: 600;
    }
    
    .modal-body {
        padding: 24px;
        max-height: 60vh;
        overflow-y: auto;
    }
    
    .close-modal {
        color: white;
        font-size: 32px;
        font-weight: bold;
        cursor: pointer;
        background: none;
        border: none;
        padding: 0;
        width: 32px;
        height: 32px;
        line-height: 32px;
        text-align: center;
        transition: all 0.2s ease;
    }
    
    .close-modal:hover {
        color: #ff6b6b;
        transform: rotate(90deg);
    }
    
    .modal-info {
        background: #f8f9fa;
        padding: 16px;
        border-radius: 4px;
        margin-bottom: 20px;
        border-left: 4px solid #1a1a1a;
    }
    
    .modal-info-item {
        display: inline-block;
        margin-right: 24px;
        font-size: 14px;
    }
    
    .modal-info-label {
        font-weight: 600;
        color: #1a1a1a;
    }
    
    .modal-info-value {
        color: #666666;
        margin-left: 8px;
    }
    
    .modal-table {
        width: 100%;
        border-collapse: collapse;
        margin-top: 16px;
    }
    
    .modal-table th {
        background: #f8f9fa;
        color: #1a1a1a;
        padding: 12px;
        text-align: left;
        font-weight: 600;
        border-bottom: 2px solid #e0e0e0;
        font-size: 14px;
    }
    
    .modal-table td {
        padding: 12px;
        border-bottom: 1px solid #f0f0f0;
        color: #2d2d2d;
        font-size: 14px;
    }
    
    .modal-table tr:hover {
        background: #fafafa;
    }
    
    .loading-spinner {
        text-align: center;
        padding: 40px;
        color: #666666;
    }
    
    @keyframes fadeIn {
        from { opacity: 0; }
        to { opacity: 1; }
    }
    
    @keyframes slideDown {
        from { 
            transform: translateY(-50px);
            opacity: 0;
        }
        to { 
            transform: translateY(0);
            opacity: 1;
        }
    }
    
    /* 모바일 최적화 (768px 이하) */
    @media screen and (max-width: 768px) {
        h1 {
            font-size: 24px !important;
        }
        
        h3 {
            font-size: 12px !important;
        }
        
        .main {
            padding: 0.5rem;
        }
        
        /* 테이블 폰트 크기 축소 */
        .responsive-table-container table {
            font-size: 11px;
        }
        
        .responsive-table-container th,
        .responsive-table-container td {
            padding: 8px 4px !important;
            font-size: 11px !important;
        }
        
        .responsive-table-container th[rowspan] {
            font-size: 12px !important;
            padding: 10px 6px !important;
        }
        
        /* 범례 모바일 최적화 */
        .legend-container {
            grid-template-columns: 1fr !important;
            gap: 12px !important;
        }
        
        .legend-container > div {
            font-size: 13px !important;
        }
        
        .legend-container span[style*="width: 24px"] {
            width: 20px !important;
            height: 20px !important;
        }
        
        /* 모달 모바일 최적화 */
        .modal-content {
            width: 95%;
            margin: 10% auto;
        }
        
        .modal-header h2 {
            font-size: 16px;
        }
        
        .modal-body {
            padding: 16px;
            max-height: 70vh;
        }
        
        .modal-info {
            padding: 12px;
        }
        
        .modal-info-item {
            display: block;
            margin-right: 0;
            margin-bottom: 8px;
            font-size: 13px;
        }
        
        .modal-table th,
        .modal-table td {
            padding: 8px 6px !important;
            font-size: 12px !important;
        }
    }
    
    /* 작은 모바일 (480px 이하) */
    @media screen and (max-width: 480px) {
        h1 {
            font-size: 20px !important;
        }
        
        .responsive-table-container table {
            font-size: 10px;
        }
        
        .responsive-table-container th,
        .responsive-table-container td {
            padding: 6px 3px !important;
            font-size: 10px !important;
        }
        
        .responsive-table-container th[rowspan] {
            font-size: 11px !important;
            padding: 8px 4px !important;
        }
        
        /* 버튼 크기 조정 */
        .stButton > button,
        .stDownloadButton > button {
            padding: 12px 16px;
            font-size: 13px;
        }
        
        /* 모달 작은 모바일 최적화 */
        .modal-content {
            width: 98%;
            margin: 5% auto;
        }
        
        .modal-header {
            padding: 16px;
        }
        
        .modal-header h2 {
            font-size: 14px;
        }
        
        .close-modal {
            font-size: 28px;
        }
        
        .modal-body {
            padding: 12px;
        }
        
        .modal-info-item {
            font-size: 12px;
        }
        
        .modal-table th,
        .modal-table td {
            padding: 6px 4px !important;
            font-size: 11px !important;
        }
    }
</style>
""", unsafe_allow_html=True)

# DB 연결 (필터용 데이터 조회)
try:
    drivers_to_try = [
        '{ODBC Driver 17 for SQL Server}',
        '{ODBC Driver 18 for SQL Server}',
        '{SQL Server}',
        '{SQL Server Native Client 11.0}'
    ]
    
    driver = None
    for d in drivers_to_try:
        try:
            conn_string_base = f"DRIVER={d};SERVER={DB_CONFIG['server']};DATABASE={DB_CONFIG['base_database']};UID={DB_CONFIG['username']};PWD={DB_CONFIG['password']}"
            test_conn = pyodbc.connect(conn_string_base)
            test_conn.close()
            driver = d
            break
        except:
            continue
    
    if driver:
        # Vessels 조회
        conn_base = pyodbc.connect(conn_string_base)
        vessels_query = "SELECT id, code, name FROM vessels WHERE deleted_at IS NULL AND is_cruise_available = 1 ORDER BY name"
        df_vessels = pd.read_sql(vessels_query, conn_base)
        
        # Routes 조회
        routes_query = "SELECT id, code, description FROM routes WHERE deleted_at IS NULL ORDER BY code"
        df_routes = pd.read_sql(routes_query, conn_base)
        
        # Ports 조회 (컬럼명 확인 필요 - 일단 주석 처리)
        # ports_query = "SELECT id, name FROM ports WHERE deleted_at IS NULL ORDER BY name"
        # df_ports = pd.read_sql(ports_query, conn_base)
        df_ports = pd.DataFrame()  # 임시로 비활성화
        conn_base.close()
    else:
        st.error("❌ ODBC 드라이버를 찾을 수 없습니다.")
        df_vessels = pd.DataFrame()
        df_routes = pd.DataFrame()
        df_ports = pd.DataFrame()
except Exception as e:
    st.warning(f"⚠️ 필터 데이터 로딩 실패: {str(e)}")
    df_vessels = pd.DataFrame()
    df_routes = pd.DataFrame()
    df_ports = pd.DataFrame()

# 필터 섹션
st.markdown('<h3 style="color: #2d2d2d; font-weight: 600; font-size: 16px; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 20px;">검색 조건</h3>', unsafe_allow_html=True)
col1, col2, col3, col4 = st.columns(4)

# 항로별 포트 매핑
route_ports = {
    'BOC': ['전체', 'PUS', 'OSA'],  # 부산-오사카
    'ONC': ['전체', 'PUS', 'OSA'],  # 부산 주말 크루즈
    'KSC': ['전체', 'PUS', 'ICN'],  # 한국해협 크루즈
    'TSL': ['전체', 'PUS', 'TSM']   # 대마도
}

with col1:
    # 항로: BOC, ONC, KSC, TSL 등
    route_options = {
        'BOC': 'BOC',
        'ONC': 'ONC', 
        'KSC': 'KSC',
        'TSL': 'TSL'
    }
    selected_route_display = st.selectbox("항로", list(route_options.keys()), index=0, key="route_select")
    selected_route = route_options[selected_route_display]

with col2:
    # 선박: 항로에 따라 자동 결정 (읽기 전용)
    vessel_map = {
        'BOC': 'PSMC',
        'ONC': 'PSMC',
        'KSC': 'PSMC',
        'TSL': 'PSTL'
    }
    selected_vessel = vessel_map.get(selected_route, 'PSMC')
    st.text_input("선박", value=selected_vessel, disabled=True, key="vessel_input")

with col3:
    # 출발지: 항로에 따라 동적 변경
    origin_options = route_ports.get(selected_route, ['전체', 'PUS', 'OSA'])
    selected_origin = st.selectbox("출발지", origin_options, index=0, key="origin_select")

with col4:
    # 도착지: 항로에 따라 동적 변경
    destination_options = route_ports.get(selected_route, ['전체', 'PUS', 'OSA'])
    selected_destination = st.selectbox("도착지", destination_options, index=0, key="destination_select")

col5, col6, col7 = st.columns([2, 2, 1])

with col5:
    start_date = st.date_input("출항시작일", datetime(2025, 11, 1))

with col6:
    end_date = st.date_input("출항종료일", datetime(2025, 11, 30))

with col7:
    st.markdown('<div style="margin-top: 28px;"></div>', unsafe_allow_html=True)
    query_button = st.button("조회", type="primary")

st.markdown('<hr style="border: none; height: 1px; background: #e0e0e0; margin: 30px 0;">', unsafe_allow_html=True)

# 조회 버튼 처리
if query_button:
    
    # DB 연결
    # 여러 드라이버 버전 자동 시도
    drivers_to_try = [
        "ODBC Driver 18 for SQL Server",
        "ODBC Driver 17 for SQL Server",
        "ODBC Driver 13 for SQL Server",
        "SQL Server Native Client 11.0",
        "SQL Server",
    ]
    
    driver = None
    for d in drivers_to_try:
        if d in pyodbc.drivers():
            driver = d
            break
    
    if not driver:
        st.error("❌ SQL Server ODBC 드라이버를 찾을 수 없습니다.")
        st.info("드라이버 설치 필요: https://go.microsoft.com/fwlink/?linkid=2249004")
        st.stop()
    
    st.info(f"사용 중인 드라이버: {driver}")
    
    conn_string_base = f"""
        Driver={{{driver}}};
        Server={DB_CONFIG['server']};
        Database={DB_CONFIG['base_database']};
        UID={DB_CONFIG['username']};
        PWD={DB_CONFIG['password']};
    """
    
    conn_string_cruise = f"""
        Driver={{{driver}}};
        Server={DB_CONFIG['server']};
        Database={DB_CONFIG['cruise_database']};
        UID={DB_CONFIG['username']};
        PWD={DB_CONFIG['password']};
    """
    
    with st.spinner('데이터 조회 중...'):
        try:
            # 1. 스케줄 조회 (neohelios_base)
            conn_base = pyodbc.connect(conn_string_base)
            
            # 항로 코드를 route_id로 변환
            route_map = {
                'BOC': 1, 'ONC': 2, 'KSC': 3, 'TSL': 5
            }
            selected_route_id = route_map.get(selected_route, 1)
            
            # 출발지/도착지에 따른 direction 결정
            # 각 항로별 Outbound/Inbound 정의
            # BOC: PUS(E) ↔ OSA(W)
            # ONC: PUS 기반
            # KSC: PUS(E) ↔ ICN(W)
            # TSL: PUS(E) ↔ TSM(W)
            
            route_direction_map = {
                'BOC': {'first': 'PUS', 'second': 'OSA'},
                'ONC': {'first': 'PUS', 'second': 'OSA'},
                'KSC': {'first': 'PUS', 'second': 'ICN'},
                'TSL': {'first': 'PUS', 'second': 'TSM'}
            }
            
            directions = []
            route_ports_info = route_direction_map.get(selected_route, {'first': 'PUS', 'second': 'OSA'})
            first_port = route_ports_info['first']
            second_port = route_ports_info['second']
            
            if selected_origin == '전체' and selected_destination == '전체':
                directions = ['E', 'W']  # 양방향
            elif selected_origin == first_port and selected_destination == '전체':
                directions = ['E']  # 첫 번째 포트 출발 = Outbound
            elif selected_origin == second_port and selected_destination == '전체':
                directions = ['W']  # 두 번째 포트 출발 = Inbound
            elif selected_origin == '전체' and selected_destination == second_port:
                directions = ['E']  # 두 번째 포트 도착 = Outbound
            elif selected_origin == '전체' and selected_destination == first_port:
                directions = ['W']  # 첫 번째 포트 도착 = Inbound
            elif selected_origin == first_port and selected_destination == second_port:
                directions = ['E']  # 첫 → 두 = Outbound
            elif selected_origin == second_port and selected_destination == first_port:
                directions = ['W']  # 두 → 첫 = Inbound
            else:
                # 기타 조합
                directions = ['E', 'W']
            
            # 여러 direction에 대해 쿼리 실행
            all_schedules = []
            for direction in directions:
                schedule_query = f"""
                    SELECT 
                        cs.id AS schedule_id,
                        CONVERT(VARCHAR, cs.etd, 23) AS etd_date,
                            voy.route_id,
                            voy.direction
                    FROM coastal_schedules cs
                    LEFT JOIN proforma_schedules ps ON cs.proforma_schedule_id = ps.id
                    LEFT JOIN voyages voy ON ps.voyage_id = voy.id
                    WHERE voy.route_id = {selected_route_id}
                          AND voy.direction = '{direction}'
                      AND CAST(cs.etd AS DATE) BETWEEN '{start_date}' AND '{end_date}'
                      AND cs.deleted_at IS NULL
                      AND cs.is_cruise_available = 1
                    ORDER BY cs.etd
                """
                df_temp = pd.read_sql(schedule_query, conn_base)
                if not df_temp.empty:
                    all_schedules.append(df_temp)
            
            conn_base.close()
            
            # 모든 스케줄 통합
            if all_schedules:
                df_schedules = pd.concat(all_schedules, ignore_index=True)
            else:
                df_schedules = pd.DataFrame()
            
            # 날짜 포맷팅 (pandas에서 처리)
            df_schedules['date'] = pd.to_datetime(df_schedules['etd_date'])
            df_schedules['date_display'] = df_schedules['date'].dt.strftime('%m월 %d일')
            df_schedules['weekday'] = df_schedules['date'].dt.day_name()
            weekday_ko = {
                'Monday': '월', 'Tuesday': '화', 'Wednesday': '수', 
                'Thursday': '목', 'Friday': '금', 'Saturday': '토', 'Sunday': '일'
            }
            df_schedules['weekday'] = df_schedules['weekday'].map(weekday_ko)
            df_schedules['date'] = df_schedules['date'].dt.date
            
            if df_schedules.empty:
                st.warning("해당 기간에 스케줄이 없습니다.")
            else:
                schedule_ids = ','.join(map(str, df_schedules['schedule_id'].tolist()))
                
                # route_id 목록 가져오기 (중복 제거)
                route_ids = df_schedules['route_id'].unique().tolist()
                route_ids_str = ','.join(map(str, route_ids))
                
                # 2. 전체 객실 수 조회 (선택한 route 기준)
                conn_cruise = pyodbc.connect(conn_string_cruise)
                total_rooms_query = f"""
                    SELECT 
                        g.code AS grade,
                        COUNT(*) AS total_rooms
                    FROM rooms r
                    JOIN grades g ON r.grade_id = g.id
                    WHERE g.route_id IN ({route_ids_str})
                      AND r.deleted_at IS NULL
                      AND g.deleted_at IS NULL
                    GROUP BY g.code
                """
                df_total_rooms = pd.read_sql(total_rooms_query, conn_cruise)
                
                # 3. 예약 현황 조회 (확정, 블록 객실 수만)
                # 확정: 실제 명단(is_temporary=0, status NOT LIKE 'REFUND%')이 1개 이상 (블록 섞여도 확정!)
                # 블록: 블록만 있고 실제 명단 0개
                # 공실: 전체 객실 - 확정 - 블록
                # REFUND 상태는 취소 티켓이므로 제외!
                booking_query = f"""
                    WITH room_status AS (
                        SELECT 
                            t.departure_schedule_id,
                            t.on_boarding_room_id,
                            g.code AS grade,
                            MAX(CASE 
                                WHEN t.is_temporary = 0 
                                     AND t.status NOT LIKE 'REFUND%'
                                THEN 1 
                                ELSE 0 
                            END) AS has_confirmed,
                            MAX(CASE 
                                WHEN t.is_temporary = 1 
                                     AND t.status NOT LIKE 'REFUND%'
                                THEN 1 
                                ELSE 0 
                            END) AS has_blocked
                        FROM tickets t
                        INNER JOIN rooms r ON t.on_boarding_room_id = r.id
                        INNER JOIN grades g ON r.grade_id = g.id
                        WHERE t.departure_schedule_id IN ({schedule_ids})
                          AND t.deleted_at IS NULL
                          AND r.deleted_at IS NULL
                          AND g.deleted_at IS NULL
                          AND t.on_boarding_room_id IS NOT NULL
                          AND t.status NOT LIKE 'REFUND%'
                        GROUP BY t.departure_schedule_id, t.on_boarding_room_id, g.code
                    )
                    SELECT 
                        departure_schedule_id AS schedule_id,
                        grade,
                        COUNT(CASE WHEN has_confirmed = 1 THEN 1 END) AS confirmed_rooms,
                        COUNT(CASE WHEN has_confirmed = 0 AND has_blocked = 1 THEN 1 END) AS blocked_rooms
                    FROM room_status
                    WHERE grade IS NOT NULL
                    GROUP BY departure_schedule_id, grade
                """
                df_bookings = pd.read_sql(booking_query, conn_cruise)
                
                # 3-1. 승객 수 조회 (티켓 수 기반)
                passenger_query = f"""
                    SELECT 
                        t.departure_schedule_id AS schedule_id,
                        g.code AS grade,
                        COUNT(CASE 
                            WHEN t.is_temporary = 0 
                                 AND t.status NOT LIKE 'REFUND%'
                            THEN 1 
                        END) AS confirmed_passengers,
                        COUNT(CASE 
                            WHEN t.is_temporary = 1 
                                 AND t.status NOT LIKE 'REFUND%'
                            THEN 1 
                        END) AS blocked_passengers
                    FROM tickets t
                    INNER JOIN rooms r ON t.on_boarding_room_id = r.id
                    INNER JOIN grades g ON r.grade_id = g.id
                    WHERE t.departure_schedule_id IN ({schedule_ids})
                      AND t.deleted_at IS NULL
                      AND r.deleted_at IS NULL
                      AND g.deleted_at IS NULL
                      AND t.on_boarding_room_id IS NOT NULL
                    GROUP BY t.departure_schedule_id, g.code
                """
                df_passengers = pd.read_sql(passenger_query, conn_cruise)
                
                # 3-2. 객실별 상세 정보 조회 (모달용)
                room_details_query = f"""
                    WITH room_status AS (
                        SELECT 
                            t.departure_schedule_id,
                            t.on_boarding_room_id,
                            r.room_number,
                            g.code AS grade,
                            MAX(CASE 
                                WHEN t.is_temporary = 0 
                                     AND t.status NOT LIKE 'REFUND%'
                                THEN 1 
                                ELSE 0 
                            END) AS has_confirmed,
                            MAX(CASE 
                                WHEN t.is_temporary = 1 
                                     AND t.status NOT LIKE 'REFUND%'
                                THEN 1 
                                ELSE 0 
                            END) AS has_blocked
                        FROM tickets t
                        INNER JOIN rooms r ON t.on_boarding_room_id = r.id
                        INNER JOIN grades g ON r.grade_id = g.id
                        WHERE t.departure_schedule_id IN ({schedule_ids})
                          AND t.deleted_at IS NULL
                          AND r.deleted_at IS NULL
                          AND g.deleted_at IS NULL
                          AND t.on_boarding_room_id IS NOT NULL
                          AND t.status NOT LIKE 'REFUND%'
                        GROUP BY t.departure_schedule_id, t.on_boarding_room_id, r.room_number, g.code
                    )
                    SELECT 
                        departure_schedule_id AS schedule_id,
                        grade,
                        room_number AS room_no,
                        CASE 
                            WHEN has_confirmed = 1 THEN '확정'
                            WHEN has_blocked = 1 THEN '블록'
                        END AS status
                    FROM room_status
                    WHERE grade IS NOT NULL
                    ORDER BY schedule_id, grade, room_number
                """
                df_room_details = pd.read_sql(room_details_query, conn_cruise)
                
                # 공실 목록도 추가 (전체 객실에서 예약된 객실 제외)
                # schedule_ids를 VALUES로 직접 생성
                schedule_values = ','.join([f"({sid})" for sid in df_schedules['schedule_id'].tolist()])
                
                vacant_rooms_query = f"""
                    SELECT 
                        cs.schedule_id,
                        g.code AS grade,
                        r.room_number AS room_no
                    FROM rooms r
                    INNER JOIN grades g ON r.grade_id = g.id
                    CROSS JOIN (
                        SELECT * FROM (VALUES {schedule_values}) AS t(schedule_id)
                    ) cs
                    WHERE g.route_id IN ({route_ids_str})
                      AND r.deleted_at IS NULL
                      AND g.deleted_at IS NULL
                      AND NOT EXISTS (
                          SELECT 1 FROM tickets t
                          WHERE t.on_boarding_room_id = r.id
                            AND t.departure_schedule_id = cs.schedule_id
                            AND t.deleted_at IS NULL
                            AND t.status NOT LIKE 'REFUND%'
                      )
                    ORDER BY schedule_id, grade, room_number
                """
                df_vacant_rooms = pd.read_sql(vacant_rooms_query, conn_cruise)
                df_vacant_rooms['status'] = '공실'
                
                conn_cruise.close()
                
                # 4. 데이터 병합 및 공실 계산
                # 모든 스케줄 x 모든 등급 조합 생성
                all_combinations = []
                for _, schedule in df_schedules.iterrows():
                    for _, grade_info in df_total_rooms.iterrows():
                        all_combinations.append({
                            'schedule_id': schedule['schedule_id'],
                            'date': schedule['date'],
                            'date_display': schedule['date_display'],
                            'weekday': schedule['weekday'],
                            'grade': grade_info['grade'],
                            'total_rooms': grade_info['total_rooms']
                        })
                
                df_all = pd.DataFrame(all_combinations)
                
                # 예약 현황 병합
                df_result = df_all.merge(df_bookings, on=['schedule_id', 'grade'], how='left')
                df_result['confirmed_rooms'] = df_result['confirmed_rooms'].fillna(0).astype(int)
                df_result['blocked_rooms'] = df_result['blocked_rooms'].fillna(0).astype(int)
                df_result['total_rooms'] = df_result['total_rooms'].astype(int)
                
                # 공실 계산
                df_result['vacant_rooms'] = df_result['total_rooms'] - df_result['confirmed_rooms'] - df_result['blocked_rooms']
                df_result['vacant_rooms'] = df_result['vacant_rooms'].clip(lower=0).astype(int)
                
                # 5. 날짜별 총계 계산
                df_totals = df_result.groupby(['date', 'date_display', 'weekday']).agg({
                    'confirmed_rooms': 'sum',
                    'blocked_rooms': 'sum',
                    'vacant_rooms': 'sum'
                }).reset_index()
                df_totals['grade'] = '총계'
                
                # 6. 총계와 등급별 데이터 합치기
                df_with_totals = pd.concat([df_totals, df_result], ignore_index=True)
                
                # 7. 날짜 표시 형식
                df_with_totals['날짜'] = df_with_totals['date_display'] + ' (' + df_with_totals['weekday'] + ')'
                
                # 8. 등급 순서 정의 (총계를 먼저, 모든 등급 포함)
                grade_order = ['총계', 'BS', 'OC', 'IC', 'RS', 'GR', 'PR', 'OR', 'DA']
                existing_grades = [g for g in grade_order if g in df_with_totals['grade'].unique()]
                
                # 9. 날짜별로 한 행씩 구성 (schedule_id 포함)
                result_rows = []
                for date_val in sorted(df_with_totals['date'].unique()):
                    date_data = df_with_totals[df_with_totals['date'] == date_val]
                    row = {
                        '날짜': date_data['날짜'].iloc[0],
                        'schedule_id': date_data['schedule_id'].iloc[0],
                        'date_raw': str(date_val)
                    }
                    
                    for grade in existing_grades:
                        grade_data = date_data[date_data['grade'] == grade]
                        if not grade_data.empty:
                            row[f'{grade}_확정'] = int(grade_data['confirmed_rooms'].iloc[0])
                            row[f'{grade}_블록'] = int(grade_data['blocked_rooms'].iloc[0])
                            row[f'{grade}_공실'] = int(grade_data['vacant_rooms'].iloc[0])
                        else:
                            row[f'{grade}_확정'] = 0
                            row[f'{grade}_블록'] = 0
                            row[f'{grade}_공실'] = 0
                    
                    result_rows.append(row)
                
                # 10. DataFrame 생성
                final_df = pd.DataFrame(result_rows)
                
                # 11. 컬럼 순서 정리 (schedule_id, date_raw 유지)
                ordered_cols = ['날짜', 'schedule_id', 'date_raw']
                for grade in existing_grades:
                    ordered_cols.extend([f'{grade}_확정', f'{grade}_블록', f'{grade}_공실'])
                
                final_df = final_df[ordered_cols]
                
                # 12. 깔끔한 테이블 생성
                st.markdown(f'<div style="background: #f8f9fa; padding: 16px 24px; border-radius: 4px; color: #1a1a1a; font-weight: 600; font-size: 16px; margin: 20px 0; border-left: 4px solid #1a1a1a;">✓ {len(df_schedules)}개 스케줄 조회 완료</div>', unsafe_allow_html=True)
                
                # HTML 테이블 생성 (반응형 컨테이너로 감싸기)
                html_table = '<div class="responsive-table-container"><table style="width:100%; border-collapse: collapse; background: white; border: 1px solid #e0e0e0;">'
                
                # 헤더 1행: 등급명 - Palantir 스타일 (등급 간 구분선 추가)
                html_table += '<thead><tr><th rowspan="2" style="background: #0a0a0a; color: #ffffff; padding: 22px; border: none; border-right: 2px solid #2a2a2a; font-weight: 500; font-size: 18px; text-transform: uppercase; letter-spacing: 1px;">Date</th>'
                for idx, grade in enumerate(existing_grades):
                    if grade == '총계':
                        bg_color = '#1a1a1a'
                    else:
                        bg_color = '#0a0a0a'
                    
                    # 마지막 등급이 아니면 진한 구분선
                    is_last_grade = (idx == len(existing_grades) - 1)
                    border_right = '1px solid #2a2a2a' if is_last_grade else '2px solid #4a4a4a'
                    
                    html_table += f'<th colspan="3" style="background: {bg_color}; color: #ffffff; padding: 22px; border: none; border-right: {border_right}; font-weight: 500; font-size: 18px; text-transform: uppercase; letter-spacing: 1px;">{grade}</th>'
                html_table += '</tr>'
                
                # 헤더 2행: 확정/블록/공실 - 공실 배경 강조 + 등급 간 구분선
                html_table += '<tr>'
                for idx, grade in enumerate(existing_grades):
                    is_last_grade = (idx == len(existing_grades) - 1)
                    grade_separator = '1px solid #e0e0e0' if is_last_grade else '2px solid #d0d0d0'
                    
                    html_table += '<th style="background: #f5f5f5; color: #6b6b6b; text-align: center; padding: 16px; font-weight: 600; border: none; border-right: 1px solid #e0e0e0; border-top: 1px solid #e0e0e0; font-size: 15px; text-transform: uppercase; letter-spacing: 0.5px;">확정</th>'
                    html_table += '<th style="background: #f5f5f5; color: #6b6b6b; text-align: center; padding: 16px; font-weight: 600; border: none; border-right: 1px solid #e0e0e0; border-top: 1px solid #e0e0e0; font-size: 15px; text-transform: uppercase; letter-spacing: 0.5px;">블록</th>'
                    html_table += f'<th style="background: #fffef5; color: #6b6b6b; text-align: center; padding: 16px; font-weight: 600; border: none; border-right: {grade_separator}; border-top: 1px solid #e0e0e0; font-size: 15px; text-transform: uppercase; letter-spacing: 0.5px;">공실</th>'
                html_table += '</tr></thead>'
                
                # 선박 정보 (항로에 따라 고정)
                vessel_name = selected_vessel
                
                # 바디 - Palantir 스타일
                html_table += '<tbody>'
                for idx, row in final_df.iterrows():
                    # 교차 행 배경 - 더 subtle하게
                    row_bg = '#ffffff' if idx % 2 == 0 else '#fafafa'
                    schedule_id = row.get('schedule_id', 0)
                    date_raw = row.get('date_raw', '')
                    
                    html_table += '<tr style="border-bottom: 1px solid #efefef; transition: background 0.2s ease;">'
                    html_table += f'<td style="background: {row_bg}; color: #0a0a0a; font-weight: 500; padding: 22px; border: none; border-right: 2px solid #d0d0d0; font-size: 17px;">{row["날짜"]}</td>'
                    
                    for idx, grade in enumerate(existing_grades):
                        confirmed = int(row.get(f'{grade}_확정', 0))
                        blocked = int(row.get(f'{grade}_블록', 0))
                        vacant = int(row.get(f'{grade}_공실', 0))
                        
                        # 등급 간 구분선 (각 등급의 공실 컬럼 오른쪽)
                        is_last_grade = (idx == len(existing_grades) - 1)
                        grade_separator = '1px solid #efefef' if is_last_grade else '2px solid #d0d0d0'
                        
                        # 확정: 다크 컬러
                        html_table += f'<td style="background: {row_bg}; color: #0a0a0a; text-align: center; padding: 20px; font-weight: 600; border: none; border-right: 1px solid #efefef; font-size: 18px;">{confirmed}</td>'
                        
                        # 블록: 그레이 톤
                        html_table += f'<td style="background: {row_bg}; color: #6b6b6b; text-align: center; padding: 20px; font-weight: 500; border: none; border-right: 1px solid #efefef; font-size: 18px;">{blocked}</td>'
                        
                        # 공실: 옅은 노란색 배경, 3 미만일 때 강조, 등급 간 구분선
                        if vacant < 3 and vacant > 0:
                            # 긴급 상황 - 빨간색 강조
                            vacant_style = f'background: #fff5f5; color: #c62828; text-align: center; padding: 20px; font-weight: 700; border: none; border-right: {grade_separator}; border-left: 3px solid #ef5350; font-size: 19px;'
                        else:
                            # 일반 공실 - 옅은 노란색 배경
                            yellow_bg = '#fffef5' if row_bg == '#ffffff' else '#fffdf0'
                            vacant_style = f'background: {yellow_bg}; color: #1565c0; text-align: center; padding: 20px; font-weight: 600; border: none; border-right: {grade_separator}; font-size: 18px;'
                        
                        html_table += f'<td style="{vacant_style}">{vacant}</td>'
                    
                    html_table += '</tr>'
                html_table += '</tbody></table></div>'
                
                # ========== 승객 수 기반 테이블 생성 ==========
                # 승객 데이터 병합
                df_pass_result = df_all[['schedule_id', 'date', 'date_display', 'weekday', 'grade']].merge(
                    df_passengers, on=['schedule_id', 'grade'], how='left'
                )
                df_pass_result['confirmed_passengers'] = df_pass_result['confirmed_passengers'].fillna(0).astype(int)
                df_pass_result['blocked_passengers'] = df_pass_result['blocked_passengers'].fillna(0).astype(int)
                
                # 승객 총계 계산
                df_pass_totals = df_pass_result.groupby(['date', 'date_display', 'weekday']).agg({
                    'confirmed_passengers': 'sum',
                    'blocked_passengers': 'sum'
                }).reset_index()
                df_pass_totals['grade'] = '총계'
                df_pass_totals['schedule_id'] = df_pass_result.groupby('date')['schedule_id'].first().values
                
                # 승객 총계와 등급별 데이터 합치기
                df_pass_with_totals = pd.concat([df_pass_totals, df_pass_result], ignore_index=True)
                df_pass_with_totals['날짜'] = df_pass_with_totals['date_display'] + ' (' + df_pass_with_totals['weekday'] + ')'
                
                # 승객 날짜별로 한 행씩 구성
                pass_result_rows = []
                for date_val in sorted(df_pass_with_totals['date'].unique()):
                    date_data = df_pass_with_totals[df_pass_with_totals['date'] == date_val]
                    row = {
                        '날짜': date_data['날짜'].iloc[0],
                        'schedule_id': date_data['schedule_id'].iloc[0],
                        'date_raw': str(date_val)
                    }
                    
                    for grade in existing_grades:
                        grade_data = date_data[date_data['grade'] == grade]
                        if not grade_data.empty:
                            row[f'{grade}_확정'] = int(grade_data['confirmed_passengers'].iloc[0])
                            row[f'{grade}_블록'] = int(grade_data['blocked_passengers'].iloc[0])
                        else:
                            row[f'{grade}_확정'] = 0
                            row[f'{grade}_블록'] = 0
                    
                    pass_result_rows.append(row)
                
                final_df_passengers = pd.DataFrame(pass_result_rows)
                
                # 조회 결과를 session_state에 저장
                st.session_state.query_result = {
                    'html_table': html_table,
                    'final_df': final_df,
                    'final_df_passengers': final_df_passengers,
                    'existing_grades': existing_grades,
                    'start_date': str(start_date),
                    'end_date': str(end_date),
                    'vessel_name': vessel_name
                }
                
                st.success("조회 완료")
                
        except Exception as e:
            st.markdown(f'<div style="background: #ffebee; border-left: 3px solid #d32f2f; padding: 15px; border-radius: 4px; color: #d32f2f; font-weight: 500; margin: 20px 0;">오류 발생: {str(e)}</div>', unsafe_allow_html=True)
            st.code(str(e))

# 조회 결과 표시 (조회 버튼과 독립적으로)
if 'query_result' in st.session_state:
    result = st.session_state.query_result
    
    # 탭 생성
    tab1, tab2 = st.tabs(["객실", "승객"])
    
    with tab1:
        # 객실 테이블 렌더링
        st.markdown(result['html_table'], unsafe_allow_html=True)
        
        # 범례 - 객실용
        st.markdown("""
        <div style="margin-top: 40px; padding: 36px; background: #ffffff; border-radius: 2px; border: 1px solid #e0e0e0;">
            <div style="color: #6b6b6b; font-weight: 600; font-size: 14px; margin-bottom: 28px; text-transform: uppercase; letter-spacing: 1.5px;">범례</div>
            <div class="legend-container" style="display: grid; grid-template-columns: repeat(2, 1fr); gap: 28px;">
                <div style="display: flex; align-items: center;">
                    <span style="display: inline-block; width: 28px; height: 28px; background: #0a0a0a; border-radius: 1px; margin-right: 14px;"></span>
                    <span style="color: #0a0a0a; font-size: 16px; font-weight: 500;">확정 (실제 명단 입력 완료)</span>
                </div>
                <div style="display: flex; align-items: center;">
                    <span style="display: inline-block; width: 28px; height: 28px; background: #6b6b6b; border-radius: 1px; margin-right: 14px;"></span>
                    <span style="color: #6b6b6b; font-size: 16px; font-weight: 500;">블록 (점유만 된 상태)</span>
                </div>
                <div style="display: flex; align-items: center;">
                    <span style="display: inline-block; width: 28px; height: 28px; background: #fffef5; border: 1px solid #1565c0; border-radius: 1px; margin-right: 14px;"></span>
                    <span style="color: #1565c0; font-size: 16px; font-weight: 500;">공실 (예약 가능한 객실)</span>
                </div>
                <div style="display: flex; align-items: center;">
                    <span style="display: inline-block; width: 28px; height: 28px; background: #c62828; border-radius: 1px; margin-right: 14px;"></span>
                    <span style="color: #c62828; font-size: 16px; font-weight: 600;">긴급 (공실 3개 미만 주의)</span>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)
                
    with tab2:
        # 승객 테이블 생성
        final_df_passengers = result['final_df_passengers']
        existing_grades = result['existing_grades']
        
        # 승객 테이블 HTML 생성
        html_pass_table = '<div class="responsive-table-container"><table style="width: 100%; border-collapse: collapse; background: #ffffff; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">'
        
        # 헤더 1행: 등급명
        html_pass_table += '<thead><tr><th rowspan="2" style="background: #0a0a0a; color: #ffffff; padding: 22px; border: none; border-right: 2px solid #2a2a2a; font-weight: 500; font-size: 18px; text-transform: uppercase; letter-spacing: 1px;">Date</th>'
        for idx, grade in enumerate(existing_grades):
            if grade == '총계':
                bg_color = '#1a1a1a'
            else:
                bg_color = '#0a0a0a'
            
            is_last_grade = (idx == len(existing_grades) - 1)
            border_right = '1px solid #2a2a2a' if is_last_grade else '2px solid #4a4a4a'
            
            html_pass_table += f'<th colspan="2" style="background: {bg_color}; color: #ffffff; padding: 22px; border: none; border-right: {border_right}; font-weight: 500; font-size: 18px; text-transform: uppercase; letter-spacing: 1px;">{grade}</th>'
        html_pass_table += '</tr>'
        
        # 헤더 2행: 확정/블록
        html_pass_table += '<tr>'
        for idx, grade in enumerate(existing_grades):
            is_last_grade = (idx == len(existing_grades) - 1)
            grade_separator = '1px solid #e0e0e0' if is_last_grade else '2px solid #d0d0d0'
            
            html_pass_table += '<th style="background: #f5f5f5; color: #6b6b6b; text-align: center; padding: 16px; font-weight: 600; border: none; border-right: 1px solid #e0e0e0; border-top: 1px solid #e0e0e0; font-size: 15px; text-transform: uppercase; letter-spacing: 0.5px;">확정</th>'
            html_pass_table += f'<th style="background: #f5f5f5; color: #6b6b6b; text-align: center; padding: 16px; font-weight: 600; border: none; border-right: {grade_separator}; border-top: 1px solid #e0e0e0; font-size: 15px; text-transform: uppercase; letter-spacing: 0.5px;">블록</th>'
        html_pass_table += '</tr></thead>'
        
        # 바디
        html_pass_table += '<tbody>'
        for idx, row in final_df_passengers.iterrows():
            row_bg = '#ffffff' if idx % 2 == 0 else '#fafafa'
            
            html_pass_table += '<tr style="border-bottom: 1px solid #efefef; transition: background 0.2s ease;">'
            html_pass_table += f'<td style="background: {row_bg}; color: #0a0a0a; font-weight: 500; padding: 22px; border: none; border-right: 2px solid #d0d0d0; font-size: 17px;">{row["날짜"]}</td>'
            
            for idx_g, grade in enumerate(existing_grades):
                confirmed = int(row.get(f'{grade}_확정', 0))
                blocked = int(row.get(f'{grade}_블록', 0))
                
                is_last_grade = (idx_g == len(existing_grades) - 1)
                grade_separator = '1px solid #efefef' if is_last_grade else '2px solid #d0d0d0'
                
                # 확정
                html_pass_table += f'<td style="background: {row_bg}; color: #0a0a0a; text-align: center; padding: 20px; font-weight: 600; border: none; border-right: 1px solid #efefef; font-size: 18px;">{confirmed}</td>'
                
                # 블록
                html_pass_table += f'<td style="background: {row_bg}; color: #6b6b6b; text-align: center; padding: 20px; font-weight: 500; border: none; border-right: {grade_separator}; font-size: 18px;">{blocked}</td>'
            
            html_pass_table += '</tr>'
        html_pass_table += '</tbody></table></div>'
        
        st.markdown(html_pass_table, unsafe_allow_html=True)
        
        # 범례 - 승객용
        st.markdown("""
        <div style="margin-top: 40px; padding: 36px; background: #ffffff; border-radius: 2px; border: 1px solid #e0e0e0;">
            <div style="color: #6b6b6b; font-weight: 600; font-size: 14px; margin-bottom: 28px; text-transform: uppercase; letter-spacing: 1.5px;">범례</div>
            <div class="legend-container" style="display: grid; grid-template-columns: repeat(2, 1fr); gap: 28px;">
                <div style="display: flex; align-items: center;">
                    <span style="display: inline-block; width: 28px; height: 28px; background: #0a0a0a; border-radius: 1px; margin-right: 14px;"></span>
                    <span style="color: #0a0a0a; font-size: 16px; font-weight: 500;">확정 (실제 명단 입력 완료)</span>
                </div>
                <div style="display: flex; align-items: center;">
                    <span style="display: inline-block; width: 28px; height: 28px; background: #6b6b6b; border-radius: 1px; margin-right: 14px;"></span>
                    <span style="color: #6b6b6b; font-size: 16px; font-weight: 500;">블록 (점유만 된 상태)</span>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    # 엑셀 다운로드 (화면과 똑같은 양식)
    st.markdown('<hr style="border: none; height: 1px; background: #e0e0e0; margin: 30px 0;">', unsafe_allow_html=True)
    
    import io
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    
    final_df = result['final_df']
    passenger_final_df = result['final_df_passengers']
    existing_grades = result['existing_grades']
    start_date = result['start_date']
    end_date = result['end_date']
    
    # 엑셀 워크북 생성
    wb = Workbook()
    
    # ==================== 시트 1: 객실 ====================
    ws = wb.active
    ws.title = '객실'
    
    # 헤더 1행: Date + 등급명
    current_col = 1
    ws.cell(1, current_col, 'Date')
    ws.merge_cells(start_row=1, start_column=current_col, end_row=2, end_column=current_col)
    current_col += 1
    
    for grade in existing_grades:
        ws.cell(1, current_col, grade)
        ws.merge_cells(start_row=1, start_column=current_col, end_row=1, end_column=current_col + 2)
        current_col += 3
    
    # 헤더 2행: 확정/블록/공실
    current_col = 2
    for grade in existing_grades:
        ws.cell(2, current_col, '확정')
        ws.cell(2, current_col + 1, '블록')
        ws.cell(2, current_col + 2, '공실')
        current_col += 3
    
    # 데이터 행
    for row_idx, row in final_df.iterrows():
        excel_row = row_idx + 3
        current_col = 1
        ws.cell(excel_row, current_col, row['날짜'])
        current_col += 1
        
        for grade in existing_grades:
            ws.cell(excel_row, current_col, int(row.get(f'{grade}_확정', 0)))
            ws.cell(excel_row, current_col + 1, int(row.get(f'{grade}_블록', 0)))
            ws.cell(excel_row, current_col + 2, int(row.get(f'{grade}_공실', 0)))
            current_col += 3
    
    # 스타일링
    header_fill = PatternFill(start_color='0a0a0a', end_color='0a0a0a', fill_type='solid')
    header_font = Font(color='FFFFFF', size=12, bold=True)
    subheader_fill = PatternFill(start_color='f5f5f5', end_color='f5f5f5', fill_type='solid')
    subheader_font = Font(color='6b6b6b', size=11, bold=True)
    yellow_fill = PatternFill(start_color='fffef5', end_color='fffef5', fill_type='solid')
    
    thin_border = Border(
        left=Side(style='thin', color='e0e0e0'),
        right=Side(style='thin', color='e0e0e0'),
        top=Side(style='thin', color='e0e0e0'),
        bottom=Side(style='thin', color='e0e0e0')
    )
    
    # 헤더 스타일
    for col in range(1, ws.max_column + 1):
        ws.cell(1, col).fill = header_fill
        ws.cell(1, col).font = header_font
        ws.cell(1, col).alignment = Alignment(horizontal='center', vertical='center')
        ws.cell(1, col).border = thin_border
        
        ws.cell(2, col).fill = subheader_fill
        ws.cell(2, col).font = subheader_font
        ws.cell(2, col).alignment = Alignment(horizontal='center', vertical='center')
        ws.cell(2, col).border = thin_border
    
    # 데이터 행 스타일 + 공실 컬럼 노란색
    for row_idx in range(3, ws.max_row + 1):
        current_col = 1
        ws.cell(row_idx, current_col).alignment = Alignment(horizontal='left', vertical='center')
        ws.cell(row_idx, current_col).border = thin_border
        current_col += 1
        
        for grade in existing_grades:
            # 확정
            ws.cell(row_idx, current_col).alignment = Alignment(horizontal='center', vertical='center')
            ws.cell(row_idx, current_col).border = thin_border
            ws.cell(row_idx, current_col).font = Font(size=11, bold=True)
            
            # 블록
            ws.cell(row_idx, current_col + 1).alignment = Alignment(horizontal='center', vertical='center')
            ws.cell(row_idx, current_col + 1).border = thin_border
            ws.cell(row_idx, current_col + 1).font = Font(color='6b6b6b', size=11)
            
            # 공실 (노란색 배경)
            ws.cell(row_idx, current_col + 2).alignment = Alignment(horizontal='center', vertical='center')
            ws.cell(row_idx, current_col + 2).border = thin_border
            ws.cell(row_idx, current_col + 2).fill = yellow_fill
            ws.cell(row_idx, current_col + 2).font = Font(color='1565c0', size=11, bold=True)
            
            current_col += 3
    
    # 컬럼 너비 조정
    from openpyxl.utils import get_column_letter
    ws.column_dimensions['A'].width = 18
    for col_idx in range(2, ws.max_column + 1):
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = 10
    
    # 행 높이
    ws.row_dimensions[1].height = 25
    ws.row_dimensions[2].height = 20
    for row_idx in range(3, ws.max_row + 1):
        ws.row_dimensions[row_idx].height = 20
    
    # ==================== 시트 2: 승객 ====================
    ws2 = wb.create_sheet(title='승객')
    
    # 헤더 1행: Date + 등급명
    current_col = 1
    ws2.cell(1, current_col, 'Date')
    ws2.merge_cells(start_row=1, start_column=current_col, end_row=2, end_column=current_col)
    current_col += 1
    
    for grade in existing_grades:
        ws2.cell(1, current_col, grade)
        ws2.merge_cells(start_row=1, start_column=current_col, end_row=1, end_column=current_col + 1)
        current_col += 2
    
    # 헤더 2행: 확정/블록 (공실 없음)
    current_col = 2
    for grade in existing_grades:
        ws2.cell(2, current_col, '확정')
        ws2.cell(2, current_col + 1, '블록')
        current_col += 2
    
    # 데이터 행
    for row_idx, row in passenger_final_df.iterrows():
        excel_row = row_idx + 3
        current_col = 1
        ws2.cell(excel_row, current_col, row['날짜'])
        current_col += 1
        
        for grade in existing_grades:
            ws2.cell(excel_row, current_col, int(row.get(f'{grade}_확정', 0)))
            ws2.cell(excel_row, current_col + 1, int(row.get(f'{grade}_블록', 0)))
            current_col += 2
    
    # 스타일링 (시트 2)
    for col in range(1, ws2.max_column + 1):
        ws2.cell(1, col).fill = header_fill
        ws2.cell(1, col).font = header_font
        ws2.cell(1, col).alignment = Alignment(horizontal='center', vertical='center')
        ws2.cell(1, col).border = thin_border
        
        ws2.cell(2, col).fill = subheader_fill
        ws2.cell(2, col).font = subheader_font
        ws2.cell(2, col).alignment = Alignment(horizontal='center', vertical='center')
        ws2.cell(2, col).border = thin_border
    
    # 데이터 행 스타일
    for row_idx in range(3, ws2.max_row + 1):
        current_col = 1
        ws2.cell(row_idx, current_col).alignment = Alignment(horizontal='left', vertical='center')
        ws2.cell(row_idx, current_col).border = thin_border
        current_col += 1
        
        for grade in existing_grades:
            # 확정
            ws2.cell(row_idx, current_col).alignment = Alignment(horizontal='center', vertical='center')
            ws2.cell(row_idx, current_col).border = thin_border
            ws2.cell(row_idx, current_col).font = Font(size=11, bold=True)
            
            # 블록
            ws2.cell(row_idx, current_col + 1).alignment = Alignment(horizontal='center', vertical='center')
            ws2.cell(row_idx, current_col + 1).border = thin_border
            ws2.cell(row_idx, current_col + 1).font = Font(color='6b6b6b', size=11)
            
            current_col += 2
    
    # 컬럼 너비 조정 (시트 2)
    ws2.column_dimensions['A'].width = 18
    for col_idx in range(2, ws2.max_column + 1):
        col_letter = get_column_letter(col_idx)
        ws2.column_dimensions[col_letter].width = 10
    
    # 행 높이 (시트 2)
    ws2.row_dimensions[1].height = 25
    ws2.row_dimensions[2].height = 20
    for row_idx in range(3, ws2.max_row + 1):
        ws2.row_dimensions[row_idx].height = 20
    
    # 저장
    output = io.BytesIO()
    wb.save(output)
    excel_data = output.getvalue()
    
    st.download_button(
        label="엑셀 다운로드",
        data=excel_data,
        file_name=f"크루즈현황_객실_승객_{start_date}_{end_date}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )

st.markdown('<hr style="border: none; height: 1px; background: #e0e0e0; margin: 40px 0;">', unsafe_allow_html=True)
st.markdown('<p style="text-align: center; color: #999999; font-size: 12px;">문제가 있으면 DB 접속 정보를 확인하세요</p>', unsafe_allow_html=True)

