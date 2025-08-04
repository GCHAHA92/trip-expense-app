# jungsan.py
import pandas as pd
import re
from pathlib import Path
import sys

# ====== 고정 열 이름 ======
DATE_COL = 'Unnamed: 13'    # 출장일자
START_TIME_COL = 'Unnamed: 9'  # 시작시간(오전/오후 판별용)
END_TIME_COL = 'Unnamed: 11'   # (예비) 종료시간
DURATION_COL = '총출장시간'     # 총출장시간 '1시간22분' 형태
VEHICLE_COL = '공용차량'        # '사용' / '미사용'

# ====== 유틸 ======
def time_to_minutes(text):
    """'1시간22분' → 총 분"""
    if pd.isna(text): 
        return 0
    s = str(text)
    h = re.search(r'(\d+)시간', s)
    m = re.search(r'(\d+)분', s)
    return (int(h.group(1)) if h else 0) * 60 + (int(m.group(1)) if m else 0)

def extract_hhmm(val):
    """셀에서 HH:MM 추출 (없으면 None)"""
    if pd.isna(val): 
        return None
    m = re.search(r'(\d{1,2}):(\d{2})', str(val))
    return m.group(0) if m else None

def classify_am_pm(hhmm):
    """HH:MM → '오전'/'오후'"""
    if not hhmm:
        return '정보없음'
    m = re.match(r'(\d{1,2}):(\d{2})', str(hhmm))
    if not m:
        return '정보없음'
    return '오전' if int(m.group(1)) < 12 else '오후'

def amount_for_minutes(mins: int) -> int:
    """분→금액: <60=0, 60~239=10,000, ≥240=20,000"""
    if mins < 60:
        return 0
    elif mins < 240:
        return 10000
    else:
        return 20000

# ====== 본 계산 ======
def summarize_trip_monthly(df: pd.DataFrame):
    # 내부 반복 헤더 제거 (일자 행)
    if 'Unnamed: 13' in df.columns:
        df = df[df['Unnamed: 13'] != '일자'].copy()

    # 날짜 파생: 고정 - '출장시작'을 기준으로 날짜/월 생성
    df['출장일자'] = pd.to_datetime(df['출장시작'], errors='coerce').dt.date
    df['출장월']   = pd.to_datetime(df['출장시작'], errors='coerce').dt.month

    # 필수 컬럼 검증
    for col in ['성명', '총출장시간', '공용차량', 'Unnamed: 9']:
        if col not in df.columns:
            print(f"❌ 필수 컬럼 누락: {col}")
            return {}

    # 유틸
    def time_to_minutes(text):
        if pd.isna(text): return 0
        s = str(text)
        h = re.search(r'(\d+)시간', s)
        m = re.search(r'(\d+)분', s)
        return (int(h.group(1)) if h else 0) * 60 + (int(m.group(1)) if m else 0)

    def extract_hhmm(val):
        if pd.isna(val): return None
        m = re.search(r'(\d{1,2}):(\d{2})', str(val))
        return m.group(0) if m else None

    def classify_am_pm(hhmm):
        if not hhmm: return '정보없음'
        m = re.match(r'(\d{1,2}):(\d{2})', str(hhmm))
        if not m: return '정보없음'
        return '오전' if int(m.group(1)) < 12 else '오후'

    def amount_for_minutes(mins: int) -> int:
        if mins < 60: return 0
        elif mins < 240: return 10000
        else: return 20000

    # 파생
    df['출장시간_분']   = df['총출장시간'].apply(time_to_minutes)
    df['시작시간_text'] = df['Unnamed: 9'].apply(extract_hhmm)
    df['시작시간_구분'] = df['시작시간_text'].apply(classify_am_pm)

    # 유효행
    df_valid = df[~df['성명'].isna() & ~pd.isna(df['출장일자'])].copy()

    results = {}
    for month in range(1, 13):
        monthly_df = df_valid[df_valid['출장월'] == month]
        if monthly_df.empty:
            continue

        def calc_day(g: pd.DataFrame):
            am = g[g['시작시간_구분'] == '오전']
            pm = g[g['시작시간_구분'] == '오후']

            am_mins = int(am['출장시간_분'].sum()) if not am.empty else 0
            pm_mins = int(pm['출장시간_분'].sum()) if not pm.empty else 0

            am_amount = amount_for_minutes(am_mins)
            pm_amount = amount_for_minutes(pm_mins)

            am_vehicle = am['공용차량'].astype(str).str.strip().eq('사용').any() if not am.empty else False
            pm_vehicle = pm['공용차량'].astype(str).str.strip().eq('사용').any() if not pm.empty else False

            # 슬롯별 차량 차감(지급이 있을 때만)
            if am_amount > 0 and am_vehicle: am_amount -= 10000
            if pm_amount > 0 and pm_vehicle: pm_amount -= 10000

            am_amount = max(0, am_amount)
            pm_amount = max(0, pm_amount)

            return pd.Series({
                '오전출장건수': len(am),
                '오후출장건수': len(pm),
                '오전_분': am_mins, '오후_분': pm_mins,
                '오전_차량사용': am_vehicle, '오후_차량사용': pm_vehicle,
                '일일지급액': min(am_amount + pm_amount, 20000)
            })

        # ▶ pandas 2.2+ 에서 경고 제거: include_groups=False
        try:
            day_summary = (
                monthly_df.groupby(['성명', '출장일자'])
                .apply(calc_day, include_groups=False)
                .reset_index()
            )
        except TypeError:
            # pandas < 2.2 인 경우 fallback
            day_summary = (
                monthly_df.groupby(['성명', '출장일자'])
                .apply(calc_day)
                .reset_index()
            )

        total = (
            day_summary.groupby('성명', as_index=False)['일일지급액']
            .sum().rename(columns={'일일지급액': '총지급액'})
        )

        counts = (
            monthly_df.groupby('성명', as_index=False)
            .agg(
                오전출장횟수=('시작시간_구분', lambda x: int((x == '오전').sum())),
                오후출장횟수=('시작시간_구분', lambda x: int((x == '오후').sum())),
                공용차량사용횟수=('공용차량',     lambda x: int((x == '사용').sum()))
            )
        )

        # 4시간 이상 행 집계 (없을 수 있음 → 안전 보강)
        over4 = monthly_df[monthly_df['출장시간_분'] >= 240].copy()
        if not over4.empty:
            over4['공용차량_구분'] = over4['공용차량'].astype(str).str.strip().map(
                lambda x: '공용차량O' if x == '사용' else '공용차량X'
            )
            over4_counts = (
                over4.groupby(['성명', '공용차량_구분'])
                .size().unstack(fill_value=0).reset_index()
            )
            # 누락 컬럼 보강
            if '공용차량O' not in over4_counts.columns: over4_counts['공용차량O'] = 0
            if '공용차량X' not in over4_counts.columns: over4_counts['공용차량X'] = 0
            over4_counts = over4_counts.rename(columns={
                '공용차량O': '4시간이상(공용차량O)',
                '공용차량X': '4시간이상(공용차량X)'
            })
        else:
            over4_counts = pd.DataFrame({
                '성명': total['성명'],
                '4시간이상(공용차량O)': 0,
                '4시간이상(공용차량X)': 0
            })

        final = (
            total.merge(counts, on='성명', how='left')
                 .merge(over4_counts, on='성명', how='left')
                 .fillna(0)
        )

        # 안전 캐스팅 (없을 경우 생성)
        for col in ['4시간이상(공용차량O)', '4시간이상(공용차량X)']:
            if col not in final.columns:
                final[col] = 0
            final[col] = final[col].astype(int)

        final['총지급액'] = (
            pd.to_numeric(final['총지급액'], errors='coerce')
            .fillna(0)
            .clip(upper=280000)
            .astype(int)
            .map('{:,}'.format)
        )

        results[f"{month}월"] = final.sort_values('성명')

    return results


def main():
    if len(sys.argv) < 2:
        print("사용법: python jungsan.py <파일명.xlsx>")
        sys.exit(1)

    input_path = Path(sys.argv[1])
    if not input_path.exists():
        print(f"❌ 파일이 존재하지 않습니다: {input_path}")
        sys.exit(1)

    print("입력파일:", input_path.name)
    # 고정: 헤더 1행
    df = pd.read_excel(input_path, header=1)

    # 날짜 열 기본 검증
    if DATE_COL not in df.columns:
        print(f"❌ 필수 열이 없습니다: '{DATE_COL}'")
        sys.exit(1)
    if df[DATE_COL].isna().all():
        print(f"❌ '{DATE_COL}' 열이 모두 비어있습니다. 파일을 확인하세요.")
        sys.exit(1)

    results = summarize_trip_monthly(df)

    if not results:
        print("⚠️ 저장할 시트가 없습니다. 데이터/열 구성을 확인하세요.")
        sys.exit(1)

    output_file = input_path.parent / "출장비_요약결과_월별.xlsx"
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        for sheet, result_df in results.items():
            result_df.to_excel(writer, sheet_name=sheet, index=False)

    print(f"[완료] {output_file} 생성됨.")

if __name__ == "__main__":
    main()
