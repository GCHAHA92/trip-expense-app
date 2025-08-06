import pandas as pd
import re
from pathlib import Path
import sys

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
    hour = int(m.group(1))
    minute = int(m.group(2))
    if hour < 12: return '오전'
    if hour == 12 and minute == 0: return '오전'
    return '오후'

def amount_for_minutes(mins: int) -> int:
    if mins < 60: return 0
    elif mins < 240: return 10000
    else: return 20000

def summarize_trip_monthly(df: pd.DataFrame):
    df = df[df['Unnamed: 13'] != '일자'].copy()
    df['출장일자'] = pd.to_datetime(df['출장시작'], errors='coerce').dt.date
    df['출장월'] = pd.to_datetime(df['출장시작'], errors='coerce').dt.month
    df['출장시간_분'] = df['총출장시간'].apply(time_to_minutes)
    df['시작시간_text'] = df['Unnamed: 9'].apply(extract_hhmm)
    df['시작시간_구분'] = df['시작시간_text'].apply(classify_am_pm)

    df_valid = df[~df['성명'].isna() & ~pd.isna(df['출장일자'])].copy()
    results = {}

    for month in range(1, 13):
        monthly_df = df_valid[df_valid['출장월'] == month]
        if monthly_df.empty: continue

        def calc_per_trip(row):
            amount = amount_for_minutes(row['출장시간_분'])
            if amount > 0 and str(row['공용차량']).strip() == '사용':
                amount -= 10000
            return max(0, amount)

        monthly_df['건별지급액'] = monthly_df.apply(calc_per_trip, axis=1)
        daily_df = (
            monthly_df.groupby(['성명', '출장일자'])['건별지급액']
            .sum().clip(upper=20000).reset_index(name='일일지급액')
        )

        total = (
            daily_df.groupby('성명', as_index=False)['일일지급액']
            .sum().rename(columns={'일일지급액': '총지급액'})
        )

        counts = (
            monthly_df.groupby('성명', as_index=False)
            .agg(
                오전출장횟수=('시작시간_구분', lambda x: int((x == '오전').sum())),
                오후출장횟수=('시작시간_구분', lambda x: int((x == '오후').sum())),
                공용차량사용횟수=('공용차량', lambda x: int((x == '사용').sum()))
            )
        )

        over4 = monthly_df[monthly_df['출장시간_분'] >= 240].copy()
        if not over4.empty:
            over4['공용차량_구분'] = over4['공용차량'].astype(str).str.strip().map(
                lambda x: '공용차량O' if x == '사용' else '공용차량X'
            )
            over4_counts = (
                over4.groupby(['성명', '공용차량_구분'])
                .size().unstack(fill_value=0).reset_index()
            )
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

        final['총지급액'] = (
            pd.to_numeric(final['총지급액'], errors='coerce')
            .fillna(0)
            .clip(upper=280000)
            .astype(int)
            .map('{:,}'.format)
        )

        for col in ['4시간이상(공용차량O)', '4시간이상(공용차량X)']:
            if col not in final.columns:
                final[col] = 0
            final[col] = final[col].astype(int)

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
    df = pd.read_excel(input_path, header=1)

    if 'Unnamed: 13' not in df.columns or df['Unnamed: 13'].isna().all():
        print(f"❌ 'Unnamed: 13' 열이 비어있거나 없습니다.")
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