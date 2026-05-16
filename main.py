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

try:
    st.image("logo.png", width=200)
except:
    pass

st.title("🚀 품목 비교 자동화 앱")
st.info("한국아사히마시나리(주) 전용 - 데이터 정밀 분석 및 시스템 안정화 버전")

# 2. 항목 행(品番/個数) 자동 탐색 함수
def get_cleaned_df(file, sheet_name):
    try:
        raw_df = pd.read_excel(file, sheet_name=sheet_name, header=None, nrows=100)
        if raw_df.empty: return None
        
        header_row_idx = None
        for i, row in raw_df.iterrows():
            row_str = "".join([str(val) for val in row.values if pd.notna(val)]).replace(" ", "")
            # 일본어 전각/반각, 오타(個수) 등 다양한 표기 대응
            if "品番" in row_str and ("個" in row_str):
                header_row_idx = i
                break
        
        if header_row_idx is not None:
            df = pd.read_excel(file, sheet_name=sheet_name, header=header_row_idx)
            df.columns = [str(c).replace(" ", "").strip() for c in df.columns]
            return df
    except:
        return None
    return None

# 3. 분석 시작
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
                        # 수량 열 이름을 '個数'로 강제 통일 (오타 방지)
                        for c in df1.columns:
                            if "個" in str(c): df1.rename(columns={c: '個数'}, inplace=True)
                        for c in df2.columns:
                            if "個" in str(c): df2.rename(columns={c: '개수'}, inplace=True)
                            
                        if '개수' in df1.columns and '개수' in df2.columns:
                            df1_c = df1[~df1['品番'].astype(str).str.contains('-000-', na=False)].copy()
                            df2_c = df2[~df2['品番'].astype(str).str.contains('-000-', na=False)].copy()
                            
                            # 데이터 추출
                            new = df2_c[~df2_c['品番'].isin(df1_c['品番'])].copy()
                            new['변경유형'], new['개수_신규'] = '신규 추가', new['개수']
                            
                            m = pd.merge(df2_c, df1_c[['品番', '개수']], on='品番', suffixes=('_신규', '_기존'))
                            chg = m[m['개수_신규'] != m['개수_기존']].copy()
                            chg['변경유형'], chg['개수_신규'] = '개수 변경', chg['개수_신규']
                            chg['개수'] = chg['개수_기존']
                            
                            dele = df1_c[~df1_c['品番'].isin(df2_c['品番'])].copy()
                            dele['변경유형'], dele['개수_신규'] = '삭제', '-'
                            
                            final = pd.concat([new, chg, dele], ignore_index=True)
                            
                            # 💡 불필요 열 완벽 제거 및 순서 밀착 (A, B, C, D열 고정)
                            essential_cols = ['변경유형', '品番', '개수', '개수_신규']
                            # 존재하지 않는 열이 있을 경우 대비
                            final_cols = [c for c in essential_cols if c in final.columns]
                            # 품명 등 기타 정보가 있다면 뒤에 붙임
                            other_cols = [c for c in final.columns if c not in final_cols and 'Unnamed' not in str(c) and '기존' not in str(c)]
                            final = final[final_cols + other_cols]

                            # 시트 생성
                            s_name = sheet[:31]
                            final.to_excel(writer, sheet_name=s_name, index=False)
                            written_sheets_count += 1
                            
                            # 스타일 적용
                            ws = writer.sheets[s_name]
                            y_f = PatternFill(start_color="FFFF00", fill_type="solid")
                            r_f = PatternFill(start_color="FFC7CE", fill_type="solid")
                            
                            for r_idx in range(2, ws.max_row + 1):
                                tp = ws.cell(row=r_idx, column=1).value
                                if tp == '삭제':
                                    for c_idx in range(1, ws.max_column + 1): ws.cell(row=r_idx, column=c_idx).fill = r_f
                                elif tp in ['신규 추가', '개수 변경']:
                                    ws.cell(row=r_idx, column=4).fill = y_f # 4열: 개수_신규
                            
                            for col in ws.columns:
                                max_l = max(len(str(cell.value or "")) for cell in col)
                                ws.column_dimensions[col[0].column_letter].width = max(max_l + 3, 12)
                            
                            summary_results.append({'시트': sheet, '총': len(final)})

            # 💡 [결과물 보장] 결과가 없어도 안내 시트를 생성하여 저장 에러 방지
            if written_sheets_count == 0:
                pd.DataFrame({'분석 결과': ['品番 또는 개수 항목을 찾지 못했거나 일치하는 시트가 없습니다.']}).to_excel(writer, sheet_name='결과 없음', index=False)

        # 4. 화면 리포트
        if written_sheets_count > 0:
            st.success(f"✅ 분석 완료! 총 {written_sheets_count}개 시트가 처리되었습니다.")
            st.download_button("💾 결과 엑셀 다운로드", data=output.getvalue(), file_name="한국아사히마시나리_비교결과.xlsx")
        else:
            st.error("⚠️ 시트 내에서 '品番' 행을 찾지 못했습니다. 엑셀 제목 줄을 확인해 주세요.")
