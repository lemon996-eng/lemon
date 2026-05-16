import streamlit as st
import pandas as pd
import io
from PIL import Image
from openpyxl.styles import PatternFill

# 1. 앱 및 로고 설정
try:
    img = Image.open("logo.png")
    st.set_page_config(page_title="품목 비교 자동화 앱", page_icon=img, layout="wide")
except:
    st.set_page_config(page_title="품목 비교 자동화 앱", page_icon="📊", layout="wide")

col_logo, _ = st.columns([1, 4])
with col_logo:
    try:
        st.image("logo.png", width=200)
    except:
        pass

st.title("🚀 품목 비교 자동화 앱")
st.info("한국아사히마시나리(주) 전용 - 헤더 자동 인식 기능이 탑재되었습니다.")
st.write("---")

# 💡 [핵심 추가] 헤더 행(제목줄)을 자동으로 찾아주는 함수
def find_header_row(file, sheet_name):
    """品番과 個数라는 단어가 포함된 행 번호를 찾아 반환합니다."""
    # 최대 20행까지만 훑어봅니다.
    temp_df = pd.read_excel(file, sheet_name=sheet_name, nrows=20, header=None)
    for i, row in temp_df.iterrows():
        row_values = [str(val).strip() for val in row.values if pd.notna(val)]
        # '品番'과 '個数'가 모두 포함된 행을 찾으면 그 행 번호를 반환
        if '品番' in row_values and '個数' in row_values:
            return i
    return 0 # 못 찾으면 기본값 0

# 파일 업로드 UI
col1, col2 = st.columns(2)
with col1:
    file1 = st.file_uploader("첫 번째(기준: A) 파일 업로드", type=['xlsx'])
with col2:
    file2 = st.file_uploader("두 번째(비교: B) 파일 업로드", type=['xlsx'])

if file1 and file2:
    if st.button("🔍 전 시트 자동 분석 시작"):
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
                    
                    # 💡 [자동 인식 적용] 각 시트별로 헤더 위치를 스스로 찾습니다.
                    header1 = find_header_row(file1, sheet)
                    header2 = find_header_row(file2, sheet2_actual_name)
                    
                    df1 = xl1.parse(sheet, header=header1)
                    df2 = xl2.parse(sheet2_actual_name, header=header2)
                    
                    # 제목 공백 제거
                    df1.columns = [str(c).strip() for c in df1.columns]
                    df2.columns = [str(c).strip() for c in df2.columns]
                    
                    if '品番' in df1.columns and '個数' in df1.columns:
                        valid_sheet_written = True
                        
                        # 특정 패턴 제외
                        df1 = df1[~df1['品番'].astype(str).str.contains('-000-', na=False)].copy()
                        df2 = df2[~df2['品番'].astype(str).str.contains('-000-', na=False)].copy()
                        
                        # 1. 신규 / 2. 변경 / 3. 삭제 추출
                        new = df2[~df2['品番'].isin(df1['品番'])].copy()
                        new['변경유형'] = '신규 추가'
                        new['個数_신규'] = new['個数']
                        
                        merged = pd.merge(df2, df1[['品番', '個数']], on='品番', suffixes=('_신규', '_기존'))
                        changed = merged[merged['個数_신규'] != merged['個数_기존']].copy()
                        changed['변경유형'] = '개수 변경'
                        changed['個수'] = changed['個数_기존']
                        
                        deleted = df1[~df1['品番'].isin(df2['品番'])].copy()
                        deleted['변경유형'] = '삭제'
                        deleted['個数_신규'] = '-'
                        
                        final = pd.concat([new, changed, deleted], ignore_index=True)
                        
                        summary_results.append({
                            '시트명': sheet,
                            '신규추가': len(new), '개수변경': len(changed), '삭제': len(deleted), '총변동': len(final)
                        })
                        
                        if len(final) == 0:
                            pd.DataFrame(columns=['品番', '個数', '個数_신규', '변경유형']).to_excel(writer, sheet_name=sheet, index=False)
                            continue
                        
                        # 열 정리 및 저장
                        invalid_cols = [c for c in final.columns if 'Unnamed:' in str(c) or str(c).strip() == '' or str(c) == '個数_기존']
                        final = final.drop(columns=invalid_cols)
                        
                        cols = list(final.columns)
                        if '個数' in cols and '個数_신규' in cols:
                            cols.remove('個수_신규')
                            idx = cols.index('個数')
                            cols.insert(idx + 1, '個数_신규')
                            final = final[cols]
                        
                        final.to_excel(writer, sheet_name=sheet, index=False)
                        
                        # 스타일 적용 (노란색/빨간색)
                        worksheet = writer.sheets[sheet]
                        y_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
                        r_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                        
                        type_idx, new_idx = None, None
                        for cell in worksheet[1]:
                            if cell.value == '변경유형': type_idx = cell.column
                            if cell.value == '個数_신규': new_idx = cell.column
                        
                        for row in range(2, worksheet.max_row + 1):
                            ctype = worksheet.cell(row=row, column=type_idx).value if type_idx else ""
                            if ctype == '삭제':
                                for col in range(1, worksheet.max_column + 1):
                                    worksheet.cell(row=row, column=col).fill = r_fill
                            elif ctype in ['신규 추가', '개수 변경'] and new_idx:
                                worksheet.cell(row=row, column=new_idx).fill = y_fill
                        
                        for col in worksheet.columns:
                            max_len = max(len(str(cell.value or '')) for cell in col)
                            worksheet.column_dimensions[col[0].column_letter].width = max(max_len + 3, 10)
            
            if not valid_sheet_written:
                pd.DataFrame({'알림': ['시트를 찾을 수 없습니다.']}).to_excel(writer, sheet_name='안내', index=False)
        
        st.success("✅ 자동 분석 완료!")
        for res in summary_results:
            if res['총변동'] > 0:
                st.info(f"📄 **{res['시트명']}**: 변동 {res['총변동']}건 (추가 {res['신규추가']}, 변경 {res['개수변경']}, 삭제 {res['삭제']})")
        
        st.download_button(label="💾 결과 엑셀 다운로드", data=output.getvalue(), file_name="품목비교결과_자동인식.xlsx")
