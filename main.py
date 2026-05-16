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
st.info("한국아사히마시나리(주) 전용 - 데이터 밀착 정렬 및 시스템 안정화 버전")

# 2. 항목 행(品番/個数) 자동 탐색 함수
def get_cleaned_df(file, sheet_name):
    try:
        # 상단 150행까지 탐색 범위를 넓힙니다.
        raw_df = pd.read_excel(file, sheet_name=sheet_name, header=None, nrows=150)
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
            # 컬럼명 정규화 (공백 제거)
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
                    df1 = get_cleaned_df(file1, sheet)
                    df2 = get_cleaned_df(file2, xl2_sheets[sheet_s])
                    
                    # 필수 항목 확인 시 분석 진행
                    if df1 is not None and df2 is not None and '品番' in df1.columns:
                        # 수량 열(個)을 찾아서 '個数'로 이름 통일 (오타 방지)
                        for c in df1.columns:
                            if "個" in str(c): df1.rename(columns={c: '個수'}, inplace=True)
                        for c in df2.columns:
                            if "個" in str(c): df2.rename(columns={c: '個数'}, inplace=True)
                        
                        # 최종적으로 필요한 열 확인
                        if '個수' in df1.columns and '個数' in df2.columns:
                            # 변동 추출 (제외 패턴 반영)
                            df1_c = df1[~df1['品番'].astype(str).str.contains('-000-', na=False)].copy()
                            df2_c = df2[~df2['品番'].astype(str).str.contains('-000-', na=False)].copy()
                            
                            # 신규 항목
                            new = df2_c[~df2_c['品番'].isin(df1_c['品番'])].copy()
                            new['변경유형'], new['個수_신규'] = '신규 추가', new['個数']
                            
                            # 개수 변경 항목
                            m = pd.merge(df2_c, df1_c[['品番', '個수']], on='品番', suffixes=('_신규', '_기존'))
                            chg = m[m['個数_신규'] != m['個수_기존']].copy()
                            chg['변경유형'], chg['個수_신규'] = '개수 변경', chg['個数_신규']
                            chg['個수'] = chg['個수_기존'] # 원래 수량 칸에는 기존 수량 유지
                            
                            # 삭제 항목
                            dele = df1_c[~df1_c['品番'].isin(df2_c['品番'])].copy()
                            dele['변경유형'], dele['個수_신규'] = '삭제', '-'
                            
                            final = pd.concat([new, chg, dele], ignore_index=True)
                            
                            # 💡 [핵심] 열 순서 밀착 정렬 (E, H열 분리 방지)
                            # '個수_신규' 컬럼명 하나로 통일
                            essential_cols = ['변경유형', '品番', '個수', '個수_신규']
                            other_cols = [c for c in final.columns if c not in essential_cols and 'Unnamed' not in str(c) and '기존' not in str(c)]
                            final = final[essential_cols + other_cols]

                            # 엑셀 시트 생성
                            s_name = sheet[:31]
                            final.to_excel(writer, sheet_name=s_name, index=False)
                            written_sheets_count += 1
                            
                            # 스타일 적용
                            ws = writer.sheets[s_name]
                            y_f = PatternFill(start_color="FFFF00", fill_type="solid")
                            r_f = PatternFill(start_color="FFC7CE", fill_type="solid")
                            
                            for r_idx in range(2, ws.max_row + 1):
                                tp = ws.cell(row=r_idx, column=1).value # 변경유형
                                if tp == '삭제':
                                    for c_idx in range(1, ws.max_column + 1):
                                        ws.cell(row=r_idx, column=c_idx).fill = r_f
                                elif tp in ['신규 추가', '개수 변경']:
                                    ws.cell(row=r_idx, column=4).fill = y_f # 4열: 個수_신규
                            
                            # 너비 조절
                            for col in ws.columns:
                                max_l = max(len(str(cell.value or "")) for cell in col)
                                ws.column_dimensions[col[0].column_letter].width = max(max_l + 3, 12)
                            
                            summary_results.append({'시트': sheet, '총': len(final), '신규': len(new), '변경': len(chg), '삭제': len(dele)})

            # 💡 [방어] 시트가 하나도 없을 때 에러 방지용 가이드 시트 생성
            if written_sheets_count == 0:
                pd.DataFrame({'알림': ['항목명(品番, 個수)을 찾지 못했거나 일치하는 시트가 없습니다.']}).to_excel(writer, sheet_name='결과 없음', index=False)

        # 4. 결과 출력
        if written_sheets_count > 0:
            st.success(f"✅ 분석 완료! ({written_sheets_count}개 시트)")
            for r in summary_results:
                if r['총'] > 0:
                    st.info(f"📄 **{r['시트']}**: {r['총']}건 (신규 {r['신규']}, 변경 {r['변경']}, 삭제 {r['삭제']})")
                else:
                    st.write(f"⚪ **{r['시트']}**: 변동 사항 없음")
            st.download_button("💾 결과 엑셀 다운로드", data=output.getvalue(), file_name="비교결과_최종.xlsx")
        else:
            st.error("⚠️ 파일 내에서 '品番' 행을 찾지 못했습니다. 엑셀 제목 행을 다시 확인해 주세요.")
