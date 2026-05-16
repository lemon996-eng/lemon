import streamlit as st
import pandas as pd
import io
from openpyxl.styles import PatternFill  # 💡 엑셀 셀 색상을 채우기 위한 도구

st.set_page_config(page_title="엑셀 비교 분석기", layout="wide")
st.title("📊 엑셀 시트별 비교 자동화 도구")

# 시트별 헤더 위치 설정
sheet_headers = {
    '本機 ﾌｨｰﾀﾞｰ': 8,
    'ｻｯｶｰﾍｯﾄﾞリスト': 6,
    '上型クイックセットチゥス': 5,
    '移動台': 5
}

col1, col2 = st.columns(2)
with col1:
    file1 = st.file_uploader("첫 번째(기준) 파일 업로드", type=['xlsx'])
with col2:
    file2 = st.file_uploader("두 번째(비교) 파일 업로드", type=['xlsx'])

if file1 and file2:
    if st.button("🔍 전 시트 비교 시작"):
        xl1 = pd.ExcelFile(file1)
        xl2 = pd.ExcelFile(file2)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for sheet in xl1.sheet_names:
                if sheet in xl2.sheet_names:
                    header_pos = sheet_headers.get(sheet, 0)
                    df1 = xl1.parse(sheet, header=header_pos)
                    df2 = xl2.parse(sheet, header=header_pos)
                    
                    # 제목 공백 제거
                    df1.columns = [str(c).strip() for c in df1.columns]
                    df2.columns = [str(c).strip() for c in df2.columns]
                    
                    if '品番' in df1.columns and '個数' in df1.columns:
                        # 1. 신규 추가 추출
                        new = df2[~df2['品番'].isin(df1['品番'])].copy()
                        new['변경유형'] = '신규 추가'
                        new['기존개수'] = '-'
                        new['個数_신규'] = new['個数'] # 신규 추가 항목도 비교를 위해 복사
                        
                        # 2. 개수 변경 추출
                        merged = pd.merge(df2, df1[['品番', '個数']], on='品番', suffixes=('_신규', '_기존'))
                        changed = merged[merged['個数_신규'] != merged['個数_기존']].copy()
                        changed['변경유형'] = '개수 변경'
                        changed['個数'] = changed['個数_기존'] # 기존 개수를 원래 個数 열에 배치
                        changed['기존개수'] = changed['個数_기존']
                        # changed 데이터에는 이미 '個数_신규' 열이 존재함
                        
                        # 3. 데이터 합치기
                        final = pd.concat([new, changed], ignore_index=True)
                        
                        # 💡 [요청사항 1] 이름이 없는 열('Unnamed: ...' 또는 빈 문자열) 삭제
                        # 문자열에 'Unnamed'가 포함되어 있거나 공백인 열 필터링
                        invalid_cols = [c for c in final.columns if 'Unnamed:' in str(c) or str(c).strip() == '']
                        final = final.drop(columns=invalid_cols)
                        
                        # 💡 [요청사항 2] 열 순서 재배치 (個数 옆에 個数_신규 배치)
                        # 현재 데이터프레임의 모든 열 목록을 가져옵니다.
                        cols = list(final.columns)
                        if '個数' in cols and '個数_신규' in cols:
                            cols.remove('個数_신규') # 원래 위치에서 제거 후
                            idx = cols.index('個数')
                            cols.insert(idx + 1, '個数_신규') # '個数' 바로 오른쪽에 삽입
                            final = final[cols] # 정렬된 순서로 데이터 재배치
                        
                        # 4. 시트에 데이터 쓰기
                        final.to_excel(writer, sheet_name=sheet, index=False)
                        
                        # 5. 엑셀 디자인 및 스타일 적용 (openpyxl 활용)
                        worksheet = writer.sheets[sheet]
                        
                        # 노란색 채우기 스타일 정의 (ARGB 색상 코드: FFFF00은 노란색)
                        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
                        
                        # '個数_신규' 열이 몇 번째 열(알파벳)에 있는지 찾기
                        target_col_idx = None
                        for cell in worksheet[1]: # 첫 번째 줄(제목행)을 훑어봅니다
                            if cell.value == '個数_신규':
                                target_col_idx = cell.column
                                break
                        
                        # 💡 [요청사항 3] '個数_신규' 열 전체를 노란색으로 칠하기 (데이터가 있는 행만)
                        if target_col_idx:
                            for row in range(2, worksheet.max_row + 1): # 2번째 줄(데이터 첫 줄)부터 끝줄까지
                                worksheet.cell(row=row, column=target_col_idx).fill = yellow_fill
                        
                        # 열 너비 자동 조절 로직 (이전 기능 유지)
                        for col in worksheet.columns:
                            max_len = max(len(str(cell.value or '')) for cell in col)
                            col_letter = col[0].column_letter
                            worksheet.column_dimensions[col_letter].width = max(max_len + 3, 10)
                        
                        st.write(f"✅ '{sheet}' 완료 (열 정렬 및 하이라이트 포함)")
        
        st.success("모든 시트 분석 및 커스텀 스타일 정렬 완료!")
        st.download_button(label="💾 결과 파일 다운로드", data=output.getvalue(), file_name="result.xlsx")
