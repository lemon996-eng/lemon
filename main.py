import streamlit as st
import pandas as pd
import io
from PIL import Image
from openpyxl.styles import PatternFill

# 1. 앱 설정
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
st.info("한국아사히마시나리(주) 전용 - 헤더 인식 및 시스템 안정화 버전")
st.write("---")

# 💡 [함수] 제목줄을 더 똑똑하게 찾습니다.
def find_header_row(file, sheet_name):
    """'品番'과 '個数' 단어가 포함된 행을 찾습니다."""
    try:
        # 데이터가 없는 시트일 수 있으므로 예외처리 추가
        temp_df = pd.read_excel(file, sheet_name=sheet_name, nrows=40, header=None)
        if temp_df.empty:
            return None
            
        for i, row in temp_df.iterrows():
            # 행의 모든 값을 문자열로 합쳐서 검색
            row_content = "".join([str(val) for val in row.values if pd.notna(val)]).replace(" ", "")
            if "品番" in row_content and "個数" in row_content:
                return i
    except:
        return None
    return None

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
        output = io.BytesIO()
        
        # 💡 [핵심 변경] 일단 메모리에 엑셀 파일을 만들 준비를 합니다.
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            
            # 나중에 시트가 하나도 없을 경우를 대비해 '임시' 상태를 기록합니다.
            written_sheets_count = 0
            xl2_sheets_stripped = {str(s).strip(): s for s in xl2.sheet_names}
            
            for sheet in xl1.sheet_names:
                sheet_stripped = str(sheet).strip()
                
                if sheet_stripped in xl2_sheets_stripped:
                    sheet2_actual_name = xl2_sheets_stripped[sheet_stripped]
                    
                    h1 = find_header_row(file1, sheet)
                    h2 = find_header_row(file2, sheet2_actual_name)
                    
                    # 제목줄을 둘 다 찾았을 때만 분석
                    if h1 is not None and h2 is not None:
                        df1 = xl1.parse(sheet, header=h1)
                        df2 = xl2.parse(sheet2_actual_name, header=h2)
                        
                        # 제목 공백 제거 및 정규화
                        df1.columns = [str(c).replace(" ", "").strip() for c in df1.columns]
                        df2.columns = [str(c).replace(" ", "").strip() for c in df2.columns]
                        
                        if '品番' in df1.columns and '個수' in df1.columns:
                            # 변동 추출
                            df1_clean = df1[~df1['品番'].astype(str).str.contains('-000-', na=False)].copy()
                            df2_clean = df2[~df2['品番'].astype(str).str.contains('-000-', na=False)].copy()
                            
                            new = df2_clean[~df2_clean['品番'].isin(df1_clean['品番'])].copy()
                            new['변경유형'] = '신규 추가'
                            new['個数_신규'] = new['個数']
                            
                            merged = pd.merge(df2_clean, df1_clean[['品番', '個数']], on='品番', suffixes=('_신규', '_기존'))
                            changed = merged[merged['個数_신규'] != merged['個数_기존']].copy()
                            changed['변경유형'] = '개수 변경'
                            changed['個数'] = changed['個数_기존']
                            
                            deleted = df1_clean[~df1_clean['品番'].isin(df2_clean['品番'])].copy()
                            deleted['변경유형'] = '삭제'
                            deleted['個数_신규'] = '-'
                            
                            final = pd.concat([new, changed, deleted], ignore_index=True)
                            
                            # 💡 시트 생성 및 쓰기
                            final.to_excel(writer, sheet_name=sheet[:30], index=False) # 시트명 글자수 제한 대응
                            written_sheets_count += 1
                            
                            summary_results.append({
                                '시트명': sheet,
                                '신규추가': len(new), '개수변경': len(changed), '삭제': len(deleted), '총변동': len(final)
                            })
                            
                            # 서식 적용
                            ws = writer.sheets[sheet[:30]]
                            y_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
                            r_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                            
                            # 열 위치 찾기
                            t_idx, n_idx = None, None
                            for cell in ws[1]:
                                if cell.value == '변경유형': t_idx = cell.column
                                if cell.value == '個数_신규': n_idx = cell.column
                                
                            for row in range(2, ws.max_row + 1):
                                ctype = ws.cell(row=row, column=t_idx).value if t_idx else ""
                                if ctype == '삭제':
                                    for c in range(1, ws.max_column + 1): ws.cell(row=row, column=c).fill = r_fill
                                elif ctype in ['신규 추가', '개수 변경'] and n_idx:
                                    ws.cell(row=row, column=n_idx).fill = y_fill
                            
                            for col in ws.columns:
                                max_len = max(len(str(cell.value or '')) for cell in col)
                                ws.column_dimensions[col[0].column_letter].width = max(max_len + 3, 10)

            # 💡 [방어 코드] 만약 분석된 시트가 하나도 없다면 안내 시트 강제 생성
            if written_sheets_count == 0:
                pd.DataFrame({'분석 결과': ['제목줄을 찾지 못했거나 일치하는 시트가 없습니다.']}).to_excel(writer, sheet_name='분석안내', index=False)

        # 5. 화면 출력
        st.success("✅ 분석 완료!")
        if written_sheets_count == 0:
            st.error("⚠️ 시트에서 '品番'과 '個数' 제목줄을 찾지 못했습니다. 엑셀 파일의 9, 7, 6행 부근에 제목이 있는지 확인해주세요.")
        else:
            for res in summary_results:
                if res['총변동'] > 0:
                    st.info(f"📄 **{res['시트명']}**: 변동 {res['총변동']}건 (추가 {res['신규추가']}, 변경 {res['개수변경']}, 삭제 {res['삭제']})")
                else:
                    st.write(f"⚪ **{res['시트명']}**: 변동 사항 없음")
        
        st.download_button(label="💾 결과 엑셀 다운로드", data=output.getvalue(), file_name="한국아사히마시나리_분석결과.xlsx")
