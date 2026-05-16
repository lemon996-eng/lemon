import streamlit as st
import pandas as pd
import io
from PIL import Image
from openpyxl.styles import PatternFill

# 1. 앱 설정 (로고 및 제목)
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
st.info("한국아사히마시나리(주) 전용 - 항목 자동 탐색 및 시스템 안정화 버전")

# 2. 항목 행(品番/個数) 자동 탐색 함수
def get_cleaned_df(file, sheet_name):
    try:
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
        xl1 = pd.ExcelFile(file1)
        xl2 = pd.ExcelFile(file2)
        summary_results = []
        output = io.BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            xl2_sheets = {str(s).strip(): s for s in xl2.sheet_names}
            written_sheets_count = 0
            
            for sheet in xl1.sheet_names:
                sheet_stripped = str(sheet).strip()
                if sheet_stripped in xl2_sheets:
                    sheet2_actual = xl2_sheets[sheet_stripped]
                    df1 = get_cleaned_df(file1, sheet)
                    df2 = get_cleaned_df(file2, sheet2_actual)
                    
                    if df1 is not None and df2 is not None and '品番' in df1.columns and '個数' in df2.columns:
                        df1 = df1[~df1['品番'].astype(str).str.contains('-000-', na=False)].copy()
                        df2 = df2[~df2['品番'].astype(str).str.contains('-000-', na=False)].copy()
                        
                        # [수정포인트 1] 한글 '수' 대신 '個数_신규'로 변수명 통일
                        new = df2[~df2['品番'].isin(df1['品番'])].copy()
                        new['변경유형'], new['個수_신규'] = '신규 추가', new['個수'] # 이 부분을 아래에서 한꺼번에 정리합니다.
                        # 변수명에서 오타가 발생하지 않도록 명시적 지정
                        new.rename(columns={'個수_신규': '個数_신규'}, inplace=True, errors='ignore')
                        
                        m = pd.merge(df2, df1[['品番', '個수']], on='品番', suffixes=('_신규', '_기존'))
                        chg = m[m['個수_신규'] != m['個수_기존']].copy()
                        chg['변경유형'], chg['個수'] = '개수 변경', chg['個수_기존']
                        chg.rename(columns={'個수_신규': '個수_신규'}, inplace=True, errors='ignore')

                        dele = df1[~df1['品番'].isin(df2['品番'])].copy()
                        dele['변경유형'], dele['個수_신규'] = '삭제', '-'
                        
                        final = pd.concat([new, chg, dele], ignore_index=True)

                        # [수정포인트 2] 결과 합치기 전 모든 '개수' 관련 열 이름을 일본어 한자 '個数'로 강제 통합
                        final.rename(columns={'個수_신규': '個数_신규', '個수': '個数'}, inplace=True, errors='ignore')
                        
                        drop_cols = [c for c in final.columns if 'Unnamed:' in str(c) or str(c).strip() == '' or '기존' in str(c)]
                        final = final.drop(columns=drop_cols, errors='ignore')
                        
                        if '個数' in final.columns and '個数_신규' in final.columns:
                            cols = list(final.columns)
                            cols.insert(cols.index('個数') + 1, cols.pop(cols.index('個数_신규')))
                            final = final[cols]

                        sheet_name_safe = sheet[:31]
                        final.to_excel(writer, sheet_name=sheet_name_safe, index=False)
                        written_sheets_count += 1
                        
                        ws = writer.sheets[sheet_name_safe]
                        y_fill = PatternFill(start_color="FFFF00", fill_type="solid")
                        r_fill = PatternFill(start_color="FFC7CE", fill_type="solid")
                        
                        t_idx, n_idx = None, None
                        for cell in ws[1]:
                            if cell.value == '변경유형': t_idx = cell.column
                            if cell.value == '個数_신규': n_idx = cell.column
                        
                        for r_idx in range(2, ws.max_row + 1):
                            v_type = ws.cell(row=r_idx, column=t_idx).value if t_idx else ""
                            if v_type == '삭제':
                                for c_idx in range(1, ws.max_column + 1):
                                    ws.cell(row=r_idx, column=c_idx).fill = r_fill
                            elif v_type in ['신규 추가', '개수 변경'] and n_idx:
                                ws.cell(row=r_idx, column=n_idx).fill = y_fill
                        
                        for col_cells in ws.columns:
                            max_l = max(len(str(cell.value or "")) for cell in col_cells)
                            ws.column_dimensions[col_cells[0].column_letter].width = max(max_l + 3, 12)
                        
                        summary_results.append({'시트': sheet, '총': len(final), '신규': len(new), '변경': len(chg), '삭제': len(dele)})

            if written_sheets_count == 0:
                pd.DataFrame({'분석 결과': ['품번(品番)과 수량(個数) 항목을 찾지 못했거나 일치하는 시트가 없습니다.']}).to_excel(writer, sheet_name='결과 없음', index=False)

        if written_sheets_count > 0:
            st.success(f"✅ 분석 완료! ({written_sheets_count}개 시트)")
            for r in summary_results:
                if r['총'] > 0:
                    st.info(f"📄 **{r['시트']}**: {r['총']}건 (추가 {r['신규']}, 변경 {r['변경']}, 삭제 {r['삭제']})")
                else:
                    st.write(f"⚪ **{r['시트']}**: 변동 사항 없음")
            st.download_button("💾 결과 엑셀 다운로드", data=output.getvalue(), file_name="비교결과_최종.xlsx")
        else:
            st.error("⚠️ 시트 내에서 '品番'과 '個数' 항목이 적힌 행을 찾을 수 없습니다.")
