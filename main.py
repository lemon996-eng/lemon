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
st.info("한국아사히마시나리(주) 전용 - 헤더 탐색 로직이 강화되었습니다.")
st.write("---")

# 💡 [강화된 헤더 탐색 함수]
def find_header_row(file, sheet_name):
    """品番과 個数 단어가 포함된 행을 더 유연하게 찾습니다."""
    # 최대 30행까지 확인
    temp_df = pd.read_excel(file, sheet_name=sheet_name, nrows=30, header=None)
    for i, row in temp_df.iterrows():
        # 모든 셀 값을 문자열로 바꾸고 공백을 제거하여 리스트 생성
        row_str = [str(val).replace(" ", "").strip() for val in row.values if pd.notna(val)]
        
        # '品番'과 '個数' 단어가 포함되어 있는지 확인
        has_pinban = any("品番" in s for s in row_str)
        has_kosu = any("個数" in s for s in row_str)
        
        if has_pinban and has_kosu:
            return i
    return None # 못 찾으면 None 반환

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
            
            # 매칭 성공한 시트가 없을 경우를 대비해 안내용 빈 시트를 미리 예약(첫 번째에 위치)
            # 나중에 실제 시트가 써지면 이 안내 시트는 무시되거나 데이터가 채워집니다.
            
            xl2_sheets_stripped = {str(s).strip(): s for s in xl2.sheet_names}
            
            for sheet in xl1.sheet_names:
                sheet_stripped = str(sheet).strip()
                
                if sheet_stripped in xl2_sheets_stripped:
                    sheet2_actual_name = xl2_sheets_stripped[sheet_stripped]
                    
                    # 헤더 위치 자동 찾기
                    h1 = find_header_row(file1, sheet)
                    h2 = find_header_row(file2, sheet2_actual_name)
                    
                    # 둘 다 헤더를 찾았을 때만 분석 진행
                    if h1 is not None and h2 is not None:
                        df1 = xl1.parse(sheet, header=h1)
                        df2 = xl2.parse(sheet2_actual_name, header=h2)
                        
                        # 컬럼명 정규화 (공백 제거)
                        df1.columns = [str(c).replace(" ", "").strip() for c in df1.columns]
                        df2.columns = [str(c).replace(" ", "").strip() for c in df2.columns]
                        
                        if '品番' in df1.columns and '個数' in df1.columns:
                            valid_sheet_written = True
                            
                            # 제외 패턴
                            df1 = df1[~df1['品番'].astype(str).str.contains('-000-', na=False)].copy()
                            df2 = df2[~df2['品番'].astype(str).str.contains('-000-', na=False)].copy()
                            
                            # 분석 로직 (신규, 변경, 삭제)
                            new = df2[~df2['品番'].isin(df1['品番'])].copy()
                            new['변경유형'] = '신규 추가'
                            new['個数_신규'] = new['個수']
                            
                            merged = pd.merge(df2, df1[['品番', '個数']], on='品番', suffixes=('_신규', '_기존'))
                            changed = merged[merged['個数_신규'] != merged['個数_기존']].copy()
                            changed['변경유형'] = '개수 변경'
                            changed['個数'] = changed['個数_기존']
                            
                            deleted = df1[~df1['品番'].isin(df2['品番'])].copy()
                            deleted['변경유형'] = '삭제'
                            deleted['個수_신규'] = '-'
                            
                            final = pd.concat([new, changed, deleted], ignore_index=True)
                            
                            summary_results.append({
                                '시트명': sheet,
                                '신규추가': len(new), '개수변경': len(changed), '삭제': len(deleted), '총변동': len(final)
                            })
                            
                            # 열 정리 및 저장
                            invalid_cols = [c for c in final.columns if 'Unnamed:' in str(c) or str(c).strip() == '' or str(c) == '個数_기존']
                            final = final.drop(columns=invalid_cols)
                            
                            # 순서 조정
                            cols = list(final.columns)
                            if '個数' in cols and '個数_신규' in cols:
                                cols.remove('個数_신규')
                                idx = cols.index('個数')
                                cols.insert(idx + 1, '個数_신규')
                                final = final[cols]
                            
                            final.to_excel(writer, sheet_name=sheet, index=False)
                            
                            # 스타일 적용
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
            
            # 💡 [핵심 방어 코드] 유효한 시트가 하나도 없을 경우 안내 시트 강제 생성
            if not valid_sheet_written:
                pd.DataFrame({'알림': ['제목줄(品番, 個数)을 찾지 못했거나 일치하는 시트가 없습니다.']}).to_excel(writer, sheet_name='분석결과없음', index=False)
        
        st.success("✅ 분석 완료!")
        if not summary_results:
            st.error("⚠️ 시트에서 '品番'과 '個数' 제목줄을 찾지 못했습니다. 엑셀 제목을 확인해 주세요.")
        else:
            for res in summary_results:
                if res['총변동'] > 0:
                    st.info(f"📄 **{res['시트명']}**: 변동 {res['총변동']}건 (추가 {res['신규추가']}, 변경 {res['개수변경']}, 삭제 {res['삭제']})")
                else:
                    st.write(f"⚪ **{res['시트명']}**: 변동 사항 없음")
        
        st.download_button(label="💾 결과 엑셀 다운로드", data=output.getvalue(), file_name="한국아사히마시나리_비교결과.xlsx")
