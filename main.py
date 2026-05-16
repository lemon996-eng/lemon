import streamlit as st
import pandas as pd
import io
from openpyxl.styles import PatternFill

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
        
        summary_results = []
        valid_sheet_written = False
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            
            xl2_sheets_stripped = {str(s).strip(): s for s in xl2.sheet_names}
            
            for sheet in xl1.sheet_names:
                sheet_stripped = str(sheet).strip()
                
                if sheet_stripped in xl2_sheets_stripped:
                    sheet2_actual_name = xl2_sheets_stripped[sheet_stripped]
                    
                    header_pos = sheet_headers.get(sheet, 0)
                    df1 = xl1.parse(sheet, header=header_pos)
                    df2 = xl2.parse(sheet2_actual_name, header=header_pos)
                    
                    # 제목 공백 제거
                    df1.columns = [str(c).strip() for c in df1.columns]
                    df2.columns = [str(c).strip() for c in df2.columns]
                    
                    if '品番' in df1.columns and '個数' in df1.columns:
                        valid_sheet_written = True
                        
                        # '-000-'가 포함된 행 제외
                        df1 = df1[~df1['品番'].astype(str).str.contains('-000-', na=False)].copy()
                        df2 = df2[~df2['品番'].astype(str).str.contains('-000-', na=False)].copy()
                        
                        # 1. 신규 추가 추출 (B에만 있는 것)
                        new = df2[~df2['品番'].isin(df1['品番'])].copy()
                        new['변경유형'] = '신규 추가'
                        new['個数_신규'] = new['個数']
                        
                        # 2. 개수 변경 추출 (A, B 둘 다 있고 수량이 다른 것)
                        merged = pd.merge(df2, df1[['品番', '個数']], on='品番', suffixes=('_신규', '_기존'))
                        changed = merged[merged['個数_신규'] != merged['個_기존']].copy()
                        changed['변경유형'] = '개수 변경'
                        changed['個数'] = changed['個数_기존']
                        
                        # 3. 💡 [새로 추가] 삭제 추출 (A에는 있지만 B에는 없는 것)
                        deleted = df1[~df1['品番'].isin(df2['品番'])].copy()
                        deleted['변경유형'] = '삭제'
                        deleted['個数_신규'] = '-' # 삭제되었으므로 신규 개수는 - 처리
                        
                        new_count = len(new)
                        change_count = len(changed)
                        delete_count = len(deleted)
                        
                        # 4. 데이터 합치기
                        final = pd.concat([new, changed, deleted], ignore_index=True)
                        
                        summary_results.append({
                            '시트명': sheet,
                            '신규추가': new_count,
                            '개수변경': change_count,
                            '삭제': delete_count,
                            '총변동': new_count + change_count + delete_count
                        })
                        
                        if len(final) == 0:
                            pd.DataFrame(columns=['品番', '個数', '個数_신규', '변경유형']).to_excel(writer, sheet_name=sheet, index=False)
                            continue
                        
                        # 원치 않는 열 삭제
                        invalid_cols = [
                            c for c in final.columns 
                            if 'Unnamed:' in str(c) 
                            or str(c).strip() == '' 
                            or str(c) in ['기존개수', '個数_기존']
                        ]
                        final = final.drop(columns=invalid_cols)
                        
                        # 열 순서 재배치
                        cols = list(final.columns)
                        if '個数' in cols and '個数_신규' in cols:
                            cols.remove('個数_신규')
                            idx = cols.index('個数')
                            cols.insert(idx + 1, '個数_신규')
                            final = final[cols]
                        
                        # 5. 시트에 데이터 쓰기
                        final.to_excel(writer, sheet_name=sheet, index=False)
                        
                        # 6. 엑셀 디자인 및 스타일 적용
                        worksheet = writer.sheets[sheet]
                        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid") # 노란색
                        red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")   # 💡 연한 빨간색 (가독성 확보)
                        
                        # 각 열의 위치(인덱스) 확인
                        type_col_idx = None # '변경유형' 열 위치
                        new_col_idx = None  # '個数_신규' 열 위치
                        
                        for cell in worksheet[1]:
                            if cell.value == '변경유형':
                                type_col_idx = cell.column
                            if cell.value == '個数_신규':
                                new_col_idx = cell.column
                        
                        # 💡 행별 조건에 맞춰 색상 칠하기
                        for row in range(2, worksheet.max_row + 1):
                            change_type = worksheet.cell(row=row, column=type_col_idx).value if type_col_idx else ""
                            
                            if change_type == '삭제':
                                # '삭제' 항목은 그 줄 전체를 연한 빨간색으로 하이라이트
                                for col in range(1, worksheet.max_column + 1):
                                    worksheet.cell(row=row, column=col).fill = red_fill
                            elif change_type in ['신규 추가', '개수 변경']:
                                # 기존 '개수 변경', '신규 추가'는 개수_신규 열만 노란색 처리
                                if new_col_idx:
                                    worksheet.cell(row=row, column=new_col_idx).fill = yellow_fill
                        
                        # 열 너비 자동 조절
                        for col in worksheet.columns:
                            max_len = max(len(str(cell.value or '')) for cell in col)
                            col_letter = col[0].column_letter
                            worksheet.column_dimensions[col_letter].width = max(max_len + 3, 10)
            
            if not valid_sheet_written:
                pd.DataFrame({'알림': ['두 파일 간에 일치하는 시트 이름이 없거나 데이터 구조가 다릅니다. 시트 이름을 확인해 주세요.']}).to_excel(writer, sheet_name='비교불가 안내', index=False)
        
        # 화면 출력 영역
        st.success("분석 프로세스가 완료되었습니다.")
        st.subheader("📋 시트별 실시간 변동 리포트")
        
        if len(summary_results) == 0:
            st.error("⚠️ 두 엑셀 파일 간에 서로 일치하는 시트 이름을 찾지 못했습니다. 파일의 탭 이름을 다시 확인해 주세요.")
        else:
            for result in summary_results:
                if result['총변동'] > 0:
                    st.info(
                        f"📄 **{result['시트명']}** 시트 결과: "
                        f"총 **{result['총변동']}건**의 변동 사항이 발견되었습니다. "
                        f"(신규 추가: {result['신규추가']}건 / 개수 변경: {result['개수변경']}건 / 삭제: {result['삭제']}건)"
                    )
                else:
                    st.write(f"⚪ **{result['시트명']}** 시트: 일치함 (변동 사항 없음)")
                
        st.write("---")
        st.download_button(label="💾 결과 파일 다운로드", data=output.getvalue(), file_name="result.xlsx")
