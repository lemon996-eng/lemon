import streamlit as st
import pandas as pd
import io
from PIL import Image
from openpyxl.styles import PatternFill

# 1. 앱 기본 설정
try:
    img = Image.open("logo.png")
    st.set_page_config(page_title="품목 비교 자동화 앱", page_icon=img, layout="wide")
except:
    st.set_page_config(page_title="품목 비교 자동화 앱", page_icon="🚀", layout="wide")

# 로고 표시
col_logo, _ = st.columns([1, 4])
with col_logo:
    try: st.image("logo.png", width=200)
    except: pass

st.title("🚀 품목 비교 자동화 앱")
st.info("한국아사히마시나리(주) 전용 - 항목 행(品番/個数) 자동 탐색 버전")
st.write("---")

# 💡 [핵심 함수] 品番과 個数가 들어있는 행을 무조건 찾아내는 함수
def get_cleaned_df(file, sheet_name):
    """행을 하나씩 검사하며 '品番'과 '個数'가 있는 줄을 찾아 그 아래 데이터를 반환합니다."""
    # 일단 헤더 없이 전체를 읽어옵니다 (최대 100행까지)
    raw_df = pd.read_excel(file, sheet_name=sheet_name, header=None, nrows=100)
    
    header_row_idx = None
    for i, row in raw_df.iterrows():
        # 줄의 내용을 하나로 합쳐서 '品番'과 '個数' 단어가 있는지 확인
        row_str = "".join([str(val) for val in row.values if pd.notna(val)]).replace(" ", "")
        if "品番" in row_str and "個数" in row_content := row_str: # 바다코끼리 연산자 대신 안전하게
            if "品番" in row_str and "個数" in row_str:
                header_row_idx = i
                break
    
    if header_row_idx is not None:
        # 찾은 행을 제목으로 다시 읽어오기
        df = pd.read_excel(file, sheet_name=sheet_name, header=header_row_idx)
        # 컬럼명 공백 제거
        df.columns = [str(c).replace(" ", "").strip() for c in df.columns]
        return df
    return None

# 파일 업로드
col1, col2 = st.columns(2)
with col1:
    file1 = st.file_uploader("첫 번째(기준: A) 파일 업로드", type=['xlsx'])
with col2:
    file2 = st.file_uploader("두 번째(비교: B) 파일 업로드", type=['xlsx'])

if file1 and file2:
    if st.button("🔍 전 시트 자동 비교 시작"):
        xl1 = pd.ExcelFile(file1)
        xl2 = pd.ExcelFile(file2)
        
        summary_results = []
        output = io.BytesIO()
        
        # 엑셀 파일 생성 시작
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            
            # 두 번째 파일 시트 목록 전처리
            xl2_sheets_stripped = {str(s).strip(): s for s in xl2.sheet_names}
            sheets_found_count = 0
            
            for sheet in xl1.sheet_names:
                sheet_stripped = str(sheet).strip()
                
                if sheet_stripped in xl2_sheets_stripped:
                    sheet2_actual = xl2_sheets_stripped[sheet_stripped]
                    
                    # 💡 자동 탐색 로직 실행
                    df1 = get_cleaned_df(file1, sheet)
                    df2 = get_cleaned_df(file2, sheet2_actual)
                    
                    # 두 파일 모두에서 '品番', '個数' 열이 확인된 경우만 진행
                    if df1 is not None and df2 is not None and '品番' in df1.columns and '個数' in df2.columns:
                        
                        # 특정 패턴 제외 (-000-)
                        df1 = df1[~df1['品番'].astype(str).str.contains('-000-', na=False)].copy()
                        df2 = df2[~df2['品番'].astype(str).str.contains('-000-', na=False)].copy()
                        
                        # 1. 신규 / 2. 변경 / 3. 삭제 데이터 추출
                        new = df2[~df2['品番'].isin(df1['品番'])].copy()
                        new['변경유형'] = '신규 추가'
                        new['個数_신규'] = new['個数']
                        
                        merged = pd.merge(df2, df1[['品番', '個数']], on='品番', suffixes=('_신규', '_기존'))
                        changed = merged[merged['個数_신규'] != merged['個数_기존']].copy()
                        changed['변경유형'] = '개수 변경'
                        changed['個数'] = changed['個数_기존']
                        
                        deleted = df1[~df1['品番'].isin(df2['品番'])].copy()
                        deleted['변경유형'] = '삭제'
                        deleted['個数_신규'] = '-'
                        
                        final = pd.concat([new, changed, deleted], ignore_index=True)
                        
                        # 열 정리 (불필요한 열 삭제)
                        drop_cols = [c for c in final.columns if 'Unnamed:' in str(c) or str(c).strip() == '' or str(c) == '個数_기존']
                        final = final.drop(columns=drop_cols)
                        
                        # 순서 조정 (品番, 個数, 個数_신규 순)
                        cols = list(final.columns)
                        if '個数' in cols and '個数_신규' in cols:
                            cols.remove('個数_신규')
                            idx = cols.index('個수')
                            cols.insert(idx + 1, '個数_신규')
                            final = final[cols]

                        # 엑셀 시트 작성
                        safe_sheet_name = sheet[:31] # 엑셀 시트명 제한(31자) 대응
                        final.to_excel(writer, sheet_name=safe_sheet_name, index=False)
                        sheets_found_count += 1
                        
                        # 스타일 입히기
                        ws = writer.sheets[safe_sheet_name]
                        yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
                        red = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                        
                        type_idx, new_idx = None, None
                        for cell in ws[1]:
                            if cell.value == '변경유형': type_idx = cell.column
                            if cell.value == '個数_신규': new_idx = cell.column
                        
                        for row in range(2, ws.max_row + 1):
                            ctype = ws.cell(row=row, column=type_idx).value if type_idx else ""
                            if ctype == '삭제':
                                for c in range(1, ws.max_column + 1): ws.cell(row=row, column=c).fill = red
                            elif ctype in ['신규 추가', '개수 변경'] and new_idx:
                                ws.cell(row=row, column=new_idx).fill = yellow
                        
                        # 열 너비 자동 조절
                        for col in ws.columns:
                            max_len = max(len(str(cell.value or '')) for cell in col)
                            ws.column_dimensions[col[0].column_letter].width = max(max_len + 3, 12)
                        
                        summary_results.append({
                            '시트명': sheet, '신규': len(new), '변경': len(changed), '삭제': len(deleted), '총': len(final)
                        })

            # 💡 [에러 방어] 생성된 시트가 하나도 없을 경우 안내 시트 강제 생성
            if sheets_found_count == 0:
                pd.DataFrame({'결과': ['品番/個数 항목을 포함한 행을 찾지 못했거나 일치하는 시트가 없습니다.']}).to_excel(writer, sheet_name='분석 결과 없음', index=False)

        # 결과 리포트 출력
        if sheets_found_count > 0:
            st.success(f"✅ 분석 완료! 총 {sheets_found_count}개의 시트를 분석했습니다.")
            for res in summary_results:
                if res['총'] > 0:
                    st.info(f"📄 **{res['시트명']}**: 총 {res['총']}건 (추가 {res['신규']}, 변경 {res['변경']}, 삭제 {res['삭제']})")
                else:
                    st.write(f"⚪ **{res['시트명']}**: 변동 없음")
            st.download_button(label="💾 결과 엑셀 다운로드", data=output.getvalue(), file_name="한국아사히마시나리_비교결과.xlsx")
        else:
            st.error("⚠️ 시트 내에서 '品番'과 '個数' 항목이 적힌 행을 찾을 수 없습니다. 엑셀 파일 형식을 확인해주세요.")
