import streamlit as st
import pandas as pd
import io

# 1. 페이지 레이아웃 및 제목 설정
st.set_page_config(page_title="엑셀 파일 병합 도구", page_icon="📊", layout="centered")
st.title("📊 엑셀 파일 자동 병합 도구")
st.write("여러 개의 엑셀 파일을 업로드하면 첫 번째 시트의 데이터를 하나로 합쳐줍니다.")

# 2. 다중 파일 업로더 UI
uploaded_files = st.file_uploader(
    "취합할 엑셀 파일들을 선택하거나 드래그 앤 드롭 하세요 (.xlsx, .xls)", 
    accept_multiple_files=True, 
    type=['xlsx', 'xls']
)

# 3. 데이터 병합 처리
if uploaded_files:
    data_frames = []
    
    # 파일 읽기 진행 표시
    with st.spinner('업로드된 파일을 읽는 중입니다...'):
        for file in uploaded_files:
            try:
                # 각 파일의 첫 번째 시트 데이터를 DataFrame으로 로드
                df = pd.read_excel(file)
                
                # 빈 파일이 아닐 경우에만 리스트에 추가
                if not df.empty:
                    data_frames.append(df)
                else:
                    st.warning(f"⚠️ '{file.name}' 파일에 데이터가 없어 제외되었습니다.")
            except Exception as e:
                st.error(f"❌ '{file.name}' 파일을 읽는 중 오류가 발생했습니다: {e}")
    
    # 성공적으로 읽은 데이터가 있다면 병합 진행
    if data_frames:
        try:
            # 행(Row) 방향으로 데이터 이어 붙이기
            merged_df = pd.concat(data_frames, ignore_index=True)
            st.success(f"✅ 총 {len(data_frames)}개의 파일이 성공적으로 병합되었습니다!")
            
            # 4. 데이터 미리보기 (상위 10개 행)
            st.subheader("👀 데이터 미리보기 (상위 10행)")
            st.dataframe(merged_df.head(10), use_container_width=True)
            
            # 5. 다운로드를 위한 엑셀 파일 변환 (메모리 버퍼 사용)
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                merged_df.to_excel(writer, index=False)
            excel_data = output.getvalue()
            
            # 6. 다운로드 버튼 다운로드
            st.download_button(
                label="📥 병합된 엑셀 파일 다운로드 (merged_data.xlsx)",
                data=excel_data,
                file_name="merged_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"💥 데이터 병합 중 오류가 발생했습니다. 열(Column) 구조를 확인해 주세요. 오류: {e}")
else:
    st.info("💡 파일을 업로드하면 이곳에 결과와 다운로드 버튼이 나타납니다.")
