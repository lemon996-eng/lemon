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
st.info("한국아사히마시나리(주) 전용 - 데이터 정밀 통합 버전")

# 2. 항목 행(品番/個数) 자동 탐색 함수
def get_cleaned_df(file, sheet_name):
    try:
        raw_df = pd.read_excel(file, sheet_name=sheet_name, header=None, nrows=100)
        header_row_idx = None
        for i, row in raw_df.iterrows():
            row_str = "".join([str(val) for val in row.values if pd.notna(val)]).replace(" ", "")
            if "品番" in row_str and "個" in row_str:
                header_row_idx = i
                break
        
        if header_row_idx is not None:
            df = pd.read_excel(file, sheet_name=sheet_name, header=header_row_idx)
            df.columns = [str(c).replace(" ", "").strip() for c in df.columns]
            return df
    except:
        return None
    return None

# 3. 파일 업로드 및 분석
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
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            xl2_sheets = {str(s).strip(): s for s in xl2.sheet_names}
            written_sheets_count = 0
            
            for sheet in xl1.sheet_names:
                sheet_s = str(sheet).strip()
                if sheet_s in xl2_sheets:
                    df1 = get_cleaned_df(file1, sheet)
                    df2 = get_cleaned_df(file2, xl2_sheets[sheet_s])
                    
                    if df1 is not None and df2 is not None and '品番' in df1.columns:
                        # 수량 컬럼명을 '個数'로 통일 (오타 방지)
                        for c in df1.columns:
                            if "個" in str(c): df1.rename(columns={c: '個数'}, inplace=True)
                        for c in df2.columns:
                            if "個" in str(c): df2.rename(columns={c: '個수'}, inplace=True)
                            
                        if '個数' in df1.columns and '個数' in df2.columns:
                            df1 = df1[~df1['品番'].astype(str).str.contains('-000-', na=False)].copy()
                            df2 = df2[~df2['品番'].astype(str).str.contains('-000-', na=False)].copy()
                            
                            # 1. 신규 추가
                            new = df2[~df2['品番'].isin(df1['品番'])].copy()
                            new['변경유형'] = '신규 추가'
                            new['個数_신규'] = new['個数']
                            
                            # 2. 개수 변경
                            m = pd.merge(df2, df1[['品番', '개수']], on='品番', suffixes=('_신규', '_기존'))
                            chg = m[m['個数_신규'] != m['個数_기존']].copy()
                            chg['변경유형'] = '개수 변경'
                            chg['개수'] = chg['개수_기존'] # 기존 수량 유지
                            
                            # 3. 삭제
                            dele = df1[~df1['品番'].isin(df2['品番'])].copy()
                            dele['변경유형'] = '삭제'
                            dele['개수_신규'] = '-'
                            
                            final = pd.concat([new, chg, dele], ignore_index=True)
                            
                            # 💡 불필요한 중복 열 정리 (한글 '수' 포함된 열 등 삭제)
                            drop_list = [c for c in final.columns if 'Unnamed' in str(c) or '기존' in str(c) or '個수' in str(c)]
                            final = final.drop(columns=drop_list, errors='ignore')
                            
                            # 💡 열 순서 고정 (변경유형, 품번, 기존개수, 신규개수 순)
                            essential = ['변경유형', '品番', '個数', '個数_신규']
                            cols = essential + [c for c in final.columns if c not in essential]
                            final = final[cols]

                            s_name = sheet[:31]
                            final.to_excel(writer, sheet_name=s_name, index=False)
                            written_sheets_count += 1
                            
                            ws = writer.sheets[s_name]
                            y_f = PatternFill(start_color="FFFF00", fill_type="solid")
                            r_f = PatternFill(start_color="FFC7CE", fill_type="solid")
                            
                            for r_idx in range(2, ws.max_row + 1):
                                tp = ws.cell(row=r_idx, column=1).value
                                if tp == '삭제':
                                    for c_idx in range(1, ws.max_column + 1):
                                        ws.cell(row=r_idx, column=c_idx).fill = r_f
                                elif tp in ['신규 추가', '개수 변경']:
                                    ws.cell(row=r_idx, column=4).fill = y_f # D열(개수_신규) 색칠
                            
                            for col in ws.columns:
                                max_l = max(len(str(cell.value or "")) for cell in col)
                                ws.column_dimensions[col[0].column_letter].width = max(max_l + 3, 12)
                            
                            summary_results.append({'시트': sheet, '총': len(final)})

            if written_sheets_count == 0:
                pd.DataFrame({'결과': ['일치하는 데이터 없음']}).to_excel(writer, sheet_name='결과 없음', index=False)

        if written_sheets_count > 0:
            st.success(f"✅ 분석 완료!")
            st.download_button("💾 결과 엑셀 다운로드", data=output.getvalue(), file_name="품목비교_최종결과.xlsx")
