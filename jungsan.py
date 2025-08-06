# jungsan.py
import pandas as pd
import re
from pathlib import Path
import sys

# ====== 고정 열 이름 ======
DATE_COL = 'Unnamed: 13'  # 출장일자
START_TIME_COL = 'Unnamed: 9'
DURATION_COL = '총출장시간'
VEHICLE_COL = '공용차량'

# ====== 유틸 ======
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
    m = re.match(r'(\d{1,2}):(\d{2})', hhmm)
    if not m: return '정보없음'
    hour, minute = int(m.group(1)), int(m.group(2))
    return '오전' if hour < 12 else '오후'  # ✅ 12:00부터 오후

def calculate_pay(mins, used_vehicle):
    if mins < 60:
        return 0
    elif mins < 240:
        amount = 10000
    else:
        amount = 20000
    if used_vehicle and amount > 0:
        amount -= 10000
    return max(amount, 0)

# ====== 본 계산 ======
def summarize_trip_monthly(df: pd.DataFrame):
    if DATE_COL in df.columns:
        df = df[df[DATE_COL] != '일자'].copy()

    df['출장일자'] = pd.to_datetime(df['출장시작'], errors='coerce').dt.date
    df['출장월']   = pd.to_datetime(df['출장시작'], errors='coerce').dt.month

    for col in ['성명', DURATION_COL, VEHICLE_COL, START_TIME_COL]:
        if col not in df.columns:
            print(f"❌ 필수 컬럼 누락: {col}")
            return {}

    df['출장시간_분'] = df[DURATION_COL].apply(time_to_minutes)
    df['시작시간'] = df[START_TIME_COL].apply(extract_hhmm)
    df['시간대'] = df['시작시간'].apply(classify_am_pm)
    df['공용차량_사용'] = df[VEHICLE_COL].astype(str).str.strip() == '사용'
    df['건별지급액'] = df.apply(
        lambda r: calculate_pay(r['출장시간_분'], r['공용차량_사용']), axis=1
    )

    df_valid = df[~df['성명'].isna() & ~pd.isna(df['출장일자'])].copy()
    results = {}

    for month in range(1, 13):
        monthly_df = df_valid[df_valid['출장월'] == month]
        if monthly_df.empty:
            continue

        def calc_day(g):
            return min(
                g.groupby('시간대')['건별지급액'].max().sum(),
                20000
            )

        daily_summary = (
            monthly_df.groupby(['성명', '출장일자'])
            .apply(calc_day)
            .reset_index(name='일일지급액')
        )

        total = daily_summary.groupby('성명', as_index=False)['일일지급액'].sum()
        total['총지급액'] = total['일일지급액'].clip(upper=280000)
        total.drop(columns='일일지급액', inplace=True)

        counts = monthly_df.groupby('성명', as_index=False).agg(
            오전출장횟수=('시간대', lambda x: (x == '오전').sum()),
            오후출장횟수=('시간대', lambda x: (x == '오후').sum()),
            공용차량사용횟수=(VEHICLE_COL, lambda x: (x == '사용').sum())
        )

        over4 = monthly_df[monthly_df['출장시간_분'] >= 240].copy()
        if not over4.empty:
            over4['공용차량_구분'] = over4[VEHICLE_COL].astype(str).str.strip().map(
                lambda x: '공용차량O' if x == '사용' else '공용차량X'
            )
            over4_counts = (
                over4.groupby(['성명', '공용차량_구분'])
                .size().unstack(fill_value=0).reset_index()
                .rename(columns={
                    '공용차량O': '4시간이상(공용차량O)',
                    '공용차량X': '4시간이상(공용차량X)'
                })
            )
        else:
            over4_counts = pd.DataFrame({'성명': total['성명']})
            over4_counts['4시간이상(공용차량O)'] = 0
            over4_counts['4시간이상(공용차량X)'] = 0

        final = (
            total.merge(counts, on='성명', how='left')
                 .merge(over4_counts, on='성명', how='left')
                 .fillna(0)
        )

        # 안전하게 처리
        final['4시간이상(공용차량O)'] = final.get('4시간이상(공용차량O)', 0).astype(int)
        final['4시간이상(공용차량X)'] = final.get('4시간이상(공용차량X)', 0).astype(int)
        final['총지급액'] = final['총지급액'].astype(int).map('{:,}'.format)

        results[f"{month}월"] = final.sort_values('성명')

    return results

# ====== 실행 ======
def main():
    if len(sys.argv) < 2:
        print("사용법: python jungsan.py <파일명.xlsx>")
        sys.exit(1)

    input_path = Path(sys.argv[1])
    if not input_path.exists():
        print(f"❌ 파일이 존재하지 않습니다: {input_path}")
        sys.exit(1)

    print("입력파일:", input_path.name)
    df = pd.read_excel(input_path, header=1)

    if DATE_COL not in df.columns or df[DATE_COL].isna().all():
        print(f"❌ '{DATE_COL}' 열이 유효하지 않습니다.")
        sys.exit(1)

    results = summarize_trip_monthly(df)
    if not results:
        print("⚠️ 저장할 시트가 없습니다.")
        sys.exit(1)

    output_file = input_path.parent / "출장비_요약결과_월별.xlsx"
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        for sheet, result_df in results.items():
            result_df.to_excel(writer, sheet_name=sheet, index=False)

    print(f"[완료] {output_file} 생성됨.")

if __name__ == "__main__":
    main()