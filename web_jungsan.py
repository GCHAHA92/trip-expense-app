# web_jungsan_tabs.py
import streamlit as st
import pandas as pd
from io import BytesIO
from jungsan import summarize_trip_monthly  # 동일 폴더의 jungsan.py 사용

st.set_page_config(page_title="출장비 정산기 (v.20250806)", layout="wide")
st.title("🚗 출장비 월별 자동 정산기 (모두 힘내세요~~!!!!!)")

uploaded_file = st.file_uploader("📁 엑셀 파일 (.xlsx) 업로드", type=["xlsx"])

def _month_key(m: str) -> int:
    """'2월' 같은 키에서 숫자만 추출해 정렬용 정수로 변환"""
    try:
        return int(str(m).replace("월", "").strip())
    except:
        return 999

def _sum_total_amount(df_month: pd.DataFrame) -> int:
    """월별 시트의 '총지급액' 합계(문자 '150,000' → 정수 150000)"""
    s = df_month['총지급액']
    if pd.api.types.is_numeric_dtype(s):
        return int(s.sum())
    vals = pd.to_numeric(s.astype(str).str.replace(",", ""), errors="coerce").fillna(0)
    return int(vals.sum())

if uploaded_file:
    df = pd.read_excel(uploaded_file, header=1)

    with st.spinner("🔍 정산 중..."):
        results = summarize_trip_monthly(df)

    if not results:
        st.warning("❌ 분석 결과가 없습니다. 파일 구조를 확인하세요.")
        st.stop()

    # 월 키 정렬
    sorted_month_keys = sorted(results.keys(), key=_month_key)



    # ---------- 다운로드 버튼 ----------
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for month in sorted_month_keys:
            results[month].to_excel(writer, sheet_name=month, index=False)
    output.seek(0)

    st.download_button(
        label="📥 정산결과 엑셀 다운로드",
        data=output,
        file_name="출장비_요약결과_월별.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # ---------- 월별 탭 ----------
# 월별 탭 표시
    st.markdown("## 📊 월별 정산 상세")
    tabs = st.tabs(sorted_month_keys)

    for tab, month in zip(tabs, sorted_month_keys):
        with tab:
            df_month = results[month].copy()
            df_month.insert(0, "No.", range(1, len(df_month) + 1))

            # 월 총합 계산
            total_amt = (
                pd.to_numeric(df_month['총지급액'].astype(str).str.replace(",", ""), errors="coerce")
                .fillna(0)
                .sum()
            )
            st.subheader(f"{month} 정산 결과 (총 지급액: {total_amt:,.0f}원)")
            st.dataframe(df_month, use_container_width=True)

else:
    st.info("위에 엑셀 파일을 업로드하면 자동 분석됩니다.")
