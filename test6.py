import streamlit as st
import pandas as pd
import plotly.express as px
import csv
import io
import re

# ==========================================
# 0. 페이지 기본 설정
# ==========================================
st.set_page_config(page_title="Money Flow", page_icon="💸", layout="wide")

# ==========================================
# 1. 금액 표시 및 텍스트 정제 함수
# ==========================================
def money(value):
    """금액을 원 단위로 포맷팅하는 함수 (오류 방지용 통일)"""
    try:
        return f"{int(value):,.0f}원"
    except (ValueError, TypeError):
        return "0원"

def preprocess_store_name(name):
    if pd.isna(name):
        return "미확인"
    name = str(name).replace(" ", "").upper()
    name = re.sub(r'\(주\)|주식회사|\[.*?\]|\(.*\)', '', name)
    return name

# ==========================================
# 2. 안전한 데이터 로더 (Tokenizing 에러 완벽 차단)
# ==========================================
@st.cache_data
def load_data(file):
    if file.name.endswith('.xlsx'):
        return pd.read_excel(file, engine='openpyxl')
        
    # CSV 파일인 경우 안전하게 한 줄씩 읽기 (C error 차단)
    raw = file.getvalue()
    try:
        content = raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        content = raw.decode("cp949", errors="replace")

    reader = csv.reader(io.StringIO(content))
    rows = []
    header_found = False
    
    for row in reader:
        if not row:
            continue
            
        # '거래일시'가 포함된 행을 진짜 헤더로 인식
        if not header_found:
            if "거래일시" in str(row[0]).replace(" ", ""):
                header_found = True
            continue

        # 헤더 이후의 데이터 처리
        if header_found:
            # 합계 행이거나 빈 행은 제외
            if len(row) < 5 or str(row[0]).strip().startswith("출금합계"):
                continue
            rows.append(row)

    # 기본 열 이름 7개 (은행 표준)
    columns = ["거래일시", "적요", "내통장표시", "구분", "거래금액", "거래후잔액", "취급점"]
    
    # 혹시 열 개수가 모자라거나 남는 경우를 위해 크기 맞춤 처리
    processed_rows = []
    for r in rows:
        r = r[:7] # 7개 초과면 자르기
        r += [''] * (7 - len(r)) # 7개 미만이면 빈칸 채우기
        processed_rows.append(r)
        
    if not processed_rows:
        return None
        
    return pd.DataFrame(processed_rows, columns=columns)

# ==========================================
# 3. 카테고리 자동 분류
# ==========================================
def auto_categorize(name, note):
    name = str(name)
    note = str(note)
    
    # 1. 금융 이동 필터링 (총지출 제외 목적)
    if '충전' in note or '충전' in name or '동백전' in name: 
        return '분류 제외(충전)'
    if any(k in name for k in ['카뱅오픈하소안', '오픈플랫폼']): 
        return '분류 제외(본인이체/송금)'

    # 2. 소비 카테고리
    if any(k in name for k in ['스타벅스','씨유','CU','GS25','세븐일레븐','이마트24','야마벤또','배달의민족','우아한형제들','요기요','맥도날드','버거킹','카페','커피','식당','베이커리','제과','파리바게뜨','리틀프랑']): 
        return '식비'
    if any(k in name for k in ['올리브영','에이블리','다이소','번개장터','무신사','쿠팡','이마트','홈플러스','쇼핑','마트','백화점']): 
        return '쇼핑'
    if any(k in name for k in ['롯데월드','위버스','예스이십사','예스24','알라딘','영화','공연','CGV','메가박스','티켓','넷플릭스','유튜브']): 
        return '문화생활'
    if any(k in name for k in ['택시','버스','지하철','교통','코레일','KORAIL','SRT','주유','티머니']): 
        return '교통비'
    if any(k in name for k in ['병원','약국','의원','의료','치과','한의원']): 
        return '의료비'
    if any(k in name for k in ['서점','학원','교재','강의','독서실','스터디']): 
        return '교육비'
    
    return '기타'

def process_dataframe(df):
    if df is None or df.empty:
        return None
        
    processed_df = df.copy()

    # 줄바꿈 문자 제거 후 날짜 변환
    processed_df['거래일시'] = processed_df['거래일시'].astype(str).str.replace(r"[\r\n]+", " ", regex=True)
    processed_df['거래일시'] = pd.to_datetime(processed_df['거래일시'], errors='coerce')
    processed_df = processed_df.dropna(subset=['거래일시'])
    
    processed_df['월'] = processed_df['거래일시'].dt.strftime('%Y-%m')
    processed_df['일자'] = processed_df['거래일시'].dt.strftime('%Y-%m-%d')

    # 금액 숫자 변환
    processed_df['거래금액'] = processed_df['거래금액'].astype(str).str.replace(',', '').str.replace('-', '').str.replace('원', '')
    processed_df['거래금액'] = pd.to_numeric(processed_df['거래금액'], errors='coerce').fillna(0).astype(int)

    # 카테고리 적용
    processed_df['정제된_거래처'] = processed_df['내통장표시'].apply(preprocess_store_name)
    processed_df['카테고리'] = processed_df.apply(lambda row: auto_categorize(row['정제된_거래처'], row['적요']), axis=1)

    return processed_df

# ==========================================
# 4. 사용자 인터페이스 (UI) 메인
# ==========================================
st.title("💸 머니플로우(Money Flow)")
st.subheader("나만의 스마트한 소비 분석 및 예산 관리")

if 'df' not in st.session_state:
    st.session_state.df = None

# 1) 파일 업로드
st.markdown("### 1. 거래내역 파일 업로드")
uploaded_file = st.file_uploader("은행에서 다운로드한 CSV 또는 XLSX 파일을 업로드해 주세요.", type=['csv', 'xlsx'])

if uploaded_file is not None:
    if 'last_uploaded' not in st.session_state or st.session_state.last_uploaded != uploaded_file.name:
        raw_df = load_data(uploaded_file)
        if raw_df is not None:
            processed_df = process_dataframe(raw_df)
            if processed_df is not None:
                st.session_state.df = processed_df
                st.session_state.last_uploaded = uploaded_file.name
        else:
            st.error("거래내역을 추출할 수 없습니다. 정상적인 은행 거래내역 파일인지 확인해주세요.")

if st.session_state.df is not None:
    df = st.session_state.df

    # 2) 데이터 확인 및 직접 수정
    st.markdown("### 2. 분류 결과 확인 및 직접 수정")
    st.info("💡 송금 내역 등 카테고리가 잘못 지정된 경우 아래 표의 '카테고리'를 클릭하여 수정하세요.")
    
    edit_cols = ['거래일시', '내통장표시', '적요', '구분', '거래금액', '카테고리']
    editor_df = df[edit_cols].copy()
    
    category_options = ["식비", "쇼핑", "교통비", "문화생활", "의료비", "교육비", "기타", 
                        "분류 제외(충전)", "분류 제외(본인이체/송금)"]
    
    edited_df = st.data_editor(
        editor_df,
        column_config={
            "카테고리": st.column_config.SelectboxColumn("카테고리 (클릭 수정)", options=category_options, required=True),
            "거래금액": st.column_config.NumberColumn("거래금액", format="%d 원")
        },
        disabled=['거래일시', '내통장표시', '적요', '구분', '거래금액'],
        use_container_width=True,
        key="data_editor"
    )
    
    df['카테고리'] = edited_df['카테고리']
    st.divider()

    # 3) 분석할 월 선택
    month_list = sorted(df['월'].unique().tolist(), reverse=True)
    selected_month = st.selectbox("📅 분석할 월을 선택하세요", month_list)

    month_df = df[df['월'] == selected_month]
    
    # '출금'이면서 '분류 제외'로 시작하지 않는 데이터만 소비로 취급!
    expense_df = month_df[(month_df['구분'].str.contains('출금', na=False)) & (~month_df['카테고리'].str.startswith('분류 제외'))]

    # 4) 월간 예산 및 총 지출
    st.markdown(f"### 3. {selected_month} 월간 예산 및 지출 요약")
    
    col1, col2 = st.columns([1, 2])
    with col1:
        budget = st.number_input("💸 이번 달 예산 (원)", min_value=0, value=1000000, step=50000)
    
    total_expense = expense_df['거래금액'].sum() if not expense_df.empty else 0
    budget_usage_rate = (total_expense / budget) * 100 if budget > 0 else 0
    remaining_budget = budget - total_expense

    with col2:
        st.metric(label="현재 총 지출 금액", value=money(total_expense))
        if remaining_budget >= 0:
            st.success(f"✅ 남은 예산: **{money(remaining_budget)}** (사용률: {budget_usage_rate:.1f}%)")
        else:
            st.error(f"🚨 초과 금액: **{money(abs(remaining_budget))}** (사용률: {budget_usage_rate:.1f}%)")
        st.progress(min(budget_usage_rate / 100, 1.0))

    st.divider()

    if total_expense > 0:
        # 5) 카테고리별 지출 금액 및 차트
        st.markdown("### 4. 소비 카테고리 분석")
        
        category_sum = expense_df.groupby('카테고리')['거래금액'].sum().reset_index()
        category_sum = category_sum.sort_values(by='거래금액', ascending=False)
        category_sum['비율(%)'] = (category_sum['거래금액'] / total_expense * 100).round(1)
        
        col_table, col_chart = st.columns(2)
        
        with col_table:
            # 표 포맷팅
            formatted_cat = category_sum.copy()
            formatted_cat['거래금액'] = formatted_cat['거래금액'].apply(money)
            formatted_cat['비율(%)'] = formatted_cat['비율(%)'].astype(str) + " %"
            st.dataframe(formatted_cat, use_container_width=True, hide_index=True)
            
            top_category = category_sum.iloc[0]['카테고리']
            st.info(f"💡 가장 지출이 많은 항목: **[{top_category}]**")
            
            comments = {
                '식비': '외식이나 배달 음식 이용이 잦지는 않았는지 확인해 보세요.',
                '교통비': '대중교통 혜택을 활용해 지출을 줄여 보세요.',
                '쇼핑': '구매 전 꼭 필요한 소비인지 한 번 더 생각해 보세요.',
                '문화생활': '할인 혜택이나 무료 프로그램을 알뜰하게 활용해 보세요.',
                '의료비': '건강 관리가 중요하므로 예방으로 의료비 부담을 줄여 보세요.',
                '교육비': '학습에 효율적으로 관리되고 있는지 점검해 보세요.',
                '기타': '분류되지 않은 지출처를 직접 카테고리 수정하여 점검해 보세요.'
            }
            if top_category in comments:
                st.write(f"💬 코멘트: {comments[top_category]}")

        with col_chart:
            # Plotly 원그래프 (한글 깨짐 없음)
            fig_pie = px.pie(category_sum, values='거래금액', names='카테고리', title='카테고리별 지출 비율', hole=0.4)
            st.plotly_chart(fig_pie, use_container_width=True)

        # 6) 날짜별 소비 흐름 (막대 차트)
        st.markdown("### 5. 날짜별 지출 흐름")
        daily_sum = expense_df.groupby('일자')['거래금액'].sum().reset_index()
        max_day = daily_sum.loc[daily_sum['거래금액'].idxmax()]
        
        st.write(f"🔥 가장 지출이 많았던 날: **{max_day['일자']}** ({money(max_day['거래금액'])})")
        
        fig_bar = px.bar(daily_sum, x='일자', y='거래금액', text_auto='.2s', title="일자별 지출액")
        st.plotly_chart(fig_bar, use_container_width=True)

    else:
        st.info("해당 월에는 총 지출에 포함할 소비 내역이 없습니다.")

    st.divider()

    # 7) 정기 구독 항목
    st.markdown("### 6. 정기 구독(결제) 의심 항목")
    sub_keywords = ['구독', '정기결제', '월정액', '넷플릭스', '유튜브', '멜론', '디즈니', '웨이브']
    
    sub_df = expense_df[
        expense_df['적요'].astype(str).str.contains('|'.join(sub_keywords), case=False) |
        expense_df['정제된_거래처'].astype(str).str.contains('|'.join(sub_keywords), case=False)
    ]

    if not sub_df.empty:
        total_sub = sub_df['거래금액'].sum()
        st.warning(f"발견된 정기 결제 건수: **{len(sub_df)}건** (총 {money(total_sub)})")
        
        sub_display = sub_df[['거래일시', '내통장표시', '적요', '거래금액']].copy()
        sub_display['거래금액'] = sub_display['거래금액'].apply(money)
        st.dataframe(sub_display, hide_index=True)
    else:
        st.success("발견된 정기 구독 항목이 없습니다.")
