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
    st.set_page_config(page_title="품목 비교 자동화 앱", page_icon="🚀", layout="wide")

try:
    st.image("logo.png", width=200)
except:
    pass

st.title("🚀 품목 비교 자동화 앱")
st.info("한국아사히마시나리(주) 전용 - 열 순서 최적화 및 시스템 안정화 버전")

# 2. 항목 행 자동 탐색 함수
def get_cleaned_df(file, sheet_name):
    try:
        # 상단 100행을 읽어 항목명이 어디 있는지 찾습니다.
        raw_df = pd.read_excel(file, sheet_name=sheet_name, header=None, nrows=100)
        header_row_idx = None
        for i, row in raw_df.iterrows():
            row_str = "".join([str(val) for val in row.values if pd.notna(val)]).replace(" ", "")
            if "品番" in row_str and ("個数" in row_str or "個수" in row_str):
                header_row_idx = i
                break
        if header_row_idx is not None:
            df = pd.read_excel(file, sheet_name=sheet_name, header=header_row_idx)
            df.columns = [str(c).replace(" ", "").strip() for c in df.columns]
            return df
    except:
        return None
    return None

# 3. 파일 업로드 UI
col1, col2 = st.columns(2)
with col1:
    file1 = st.file_uploader("첫 번째(기준: A) 파일 업로드", type=['xlsx'])
with col2:
    file2 = st.file_uploader("두 번째(비교: B) 파일 업로드", type=['xlsx'])

if file1 and file2:
    if st.button("🔍 전 시트 자동 분석 시작"):
        xl1, xl2 = pd.ExcelFile(file1), pd.ExcelFile(file2)
        summary_results = []
        output = io.BytesIO()
        
        # 엑셀 쓰기 엔진 시작
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            xl2_sheets = {str(s).strip(): s for s in xl2.sheet_names}
            written_sheets_count = 0
            
            for sheet in xl1.sheet_names:
                sheet_s = str(sheet).strip()
                if sheet_s in xl2_sheets:
                    df1, df2 = get_cleaned_df(file1, sheet), get_cleaned_df(file2, xl2_sheets[sheet_s])
                    
                    if df1 is not None and df2 is not None and '品番' in df1.columns and '個数' in df2.columns:
                        # 제외 패턴 처리
                        df1 = df1[~df1['品番'].astype(str).str.contains('-000-', na=False)].copy()
                        df2 = df2[~df2['品番'].astype(str).str.contains('-000-', na=False)].copy()
                        
                        # 데이터 추출 (신규/변경/삭제)
                        new = df2[~df2['品番'].isin(df1['品番'])].copy()
                        new['변경유형'], new['個数_신규'] = '신규 추가', new['個数']
                        
                        m = pd.merge(df2, df1[['品番', '個数']], on='品番', suffixes=('_신규', '_기존'))
                        chg = m[m['個数_신규'] != m['個수_기존']].copy()
                        chg['변경유형'], chg['個数'] = '개수 변경', chg['個数_기존']
                        
                        dele = df1[~df1['品番'].isin(df2['品番'])].copy()
                        dele['변경유형'], dele['個数_신규'] = '삭제', '-'
                        
                        final = pd.concat([new, chg, dele], ignore_index=True)
                        
                        # 💡 [핵심] 열 순서 강제 고정 (E, H열 분리 방지)
                        # 필수 열 4개를 맨 앞으로 배치
                        essential_cols = ['변경유형', '品番', '個数', '個数_신규']
                        # 나머지 열(품명 등) 자동 감지
                        other_cols = [c for c in final.columns if c not in essential_cols and 'Unnamed' not in str(c) and '기존' not in str(c)]
                        final = final[essential_cols + other_cols]

                        # 엑셀 시트 생성
                        sheet_name_safe = sheet[:31]
                        final.to_excel(writer, sheet_name=sheet_name_safe, index=False)
                        written_sheets_count += 1
                        
                        # 스타일 적용
                        ws = writer.sheets[sheet_name_safe]
                        y_f = PatternFill(start_color="FFFF00", fill_type="solid")
                        r_f = PatternFill(start_color="FFC7CE", fill_type="solid")
                        
                        # 열 위치(1부터 시작)
                        t_idx = final.columns.get_loc('변경유형') + 1
                        n_idx = final.columns.get_loc('個数_신규') + 1
                        
                        for r in range(2, ws.max_row + 1):
                            tp = ws.cell(row=r, column=t_idx).value
                            if tp == '삭제':
                                for c in range(1, ws.max_column + 1): ws.cell(row=r, column=c).fill = r_f
                            elif tp in ['신규 추가', '개수 변경']:
                                ws.cell(row=r, column=n_idx).fill = y_f
                        
                        # 열 너비 자동 조절
                        for col_cells in ws.columns:
                            max_l = max(len(str(cell.value or "")) for cell in col_cells)
                            ws.column_dimensions[col_cells[0].column_letter].width = max(max_l + 3, 12)
                        
                        summary_results.append({'시트': sheet, '총': len(final)})

            # 💡 [방어 코드] 시트가 하나도 없을 때 에러 방지용 시트 생성
            if written_sheets_count == 0:
                pd.DataFrame({'분석 결과': ['品番/個数 항목을 찾지 못했거나 일치하는 시트가 없습니다.']}).to_excel(writer, sheet_name='결과 없음', index=False)

        # 4. 결과 출력
        if written_sheets_count > 0:
            st.success(f"✅ 분석 완료! ({written_sheets_count}개 시트)")
            st.download_button("💾 결과 엑셀 다운로드", data=output.getvalue(), file_name="비교결과_최종.xlsx")
        else:
            st.error("⚠️ '品番'과 '個数' 행을 찾지 못했습니다. 엑셀의 항목명 오타나 위치를 확인해주세요.")
