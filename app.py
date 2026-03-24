import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.chart.data import CategoryChartData
import io

# 1. 구글 시트 CSV 게시 링크
SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRhvBRSiCXgsdh9ipTVohLneYp55ROUHbSpOKUsQlVOJeuSB5bXbG7PDRWaB2VzrF57j2QPuy0G7Bcx/pub?output=csv"

# 2. 로컬에 저장된 PPT 양식 파일명
PPTX_TEMPLATE_PATH = "우리 팀의 오늘을 묻다_양식.pptx"

@st.cache_data(ttl=60)
def load_data(url):
    return pd.read_csv(url)

def update_chart_data(chart, new_values):
    """기존 차트의 X축(Category)을 유지하면서 데이터(Value)만 교체합니다."""
    chart_data = CategoryChartData()
    
    categories = [c.label for c in chart.plots[0].categories]
    chart_data.categories = categories
    
    series_name = chart.series[0].name if chart.series else "Series 1"
    chart_data.add_series(series_name, new_values)
    
    chart.replace_data(chart_data)

def remove_unused_textboxes(slide, q1_score):
    """1번 문항 점수에 따라 해당하는 텍스트 상자 1개만 남기고 나머지를 삭제합니다."""
    target_score = max(1, min(4, int(round(q1_score))))
    emojis = {1: "🌧️", 2: "☁️", 3: "⛅", 4: "🌞"}
    target_emoji = emojis[target_score]
    
    shapes_to_delete = []
    seen_ids = set() # 중복 방지를 위해 도형의 고유 ID를 저장
    
    for shape in slide.shapes:
        shape_id = shape.shape_id
        
        if shape.has_text_frame:
            text = shape.text.strip()
            if text in emojis.values():
                if text != target_emoji and shape_id not in seen_ids:
                    shapes_to_delete.append(shape)
                    seen_ids.add(shape_id)
                continue
                
        name = shape.name.lower().replace(" ", "")
        if "textbox1" in name or "텍스트상자1" in name:
            if target_score != 1 and shape_id not in seen_ids:
                shapes_to_delete.append(shape)
                seen_ids.add(shape_id)
        elif "textbox2" in name or "텍스트상자2" in name:
            if target_score != 2 and shape_id not in seen_ids:
                shapes_to_delete.append(shape)
                seen_ids.add(shape_id)
        elif "textbox3" in name or "텍스트상자3" in name:
            if target_score != 3 and shape_id not in seen_ids:
                shapes_to_delete.append(shape)
                seen_ids.add(shape_id)
        elif "textbox4" in name or "텍스트상자4" in name:
            if target_score != 4 and shape_id not in seen_ids:
                shapes_to_delete.append(shape)
                seen_ids.add(shape_id)

    # 리스트에 수집된 도형 삭제
    for shape in shapes_to_delete:
        try:
            sp = shape._element
            sp.getparent().remove(sp)
        except Exception:
            pass
        try:
            sp = shape._element
            sp.getparent().remove(sp)
        except Exception:
            pass

def main():
    st.title("설문 결과 PPT 자동 생성 툴")
    
    col1, col2 = st.columns([4, 1])
    with col1:
        st.write("원하는 차수(날짜)를 선택하고 PPT를 다운로드하세요.")
    with col2:
        # 데이터 새로고침 버튼
        if st.button("🔄 데이터 새로고침"):
            load_data.clear() # 캐시 초기화
            st.rerun() # 앱 재실행

    try:
        df = load_data(SHEET_CSV_URL)
    except Exception as e:
        st.error("구글 시트 데이터를 불러오는 데 실패했습니다. 링크를 확인해 주세요.")
        return

    date_column = '날짜' 
    if date_column not in df.columns:
        st.error(f"데이터에 '{date_column}' 컬럼이 존재하지 않습니다.")
        return
        
    dates = df[date_column].dropna().unique()
    selected_date = st.selectbox("출력할 날짜(차수)를 선택하세요:", sorted(dates))
    
    if selected_date:
        filtered_df = df[df[date_column] == selected_date]
        
        q_cols = [col for col in df.columns if col.startswith(('1.', '2.', '3.', '4.', '5.', '6.', '7.', '8.', '9.'))]
        
        if len(q_cols) < 9:
            st.error("1번부터 9번까지의 문항 컬럼을 모두 찾을 수 없습니다. 시트의 열 제목을 확인해 주세요.")
            return
            
        averages = filtered_df[q_cols].mean().round(1).tolist()
        q1_score = averages[0]
        
        st.write(f"**{selected_date}** 평균 데이터 미리보기:")
        st.dataframe(pd.DataFrame([averages], columns=[f"Q{i+1}" for i in range(9)]))

        if st.button("PPT 생성 및 다운로드 준비"):
            try:
                prs = Presentation(PPTX_TEMPLATE_PATH)
            except Exception as e:
                st.error(f"PPT 양식 파일을 찾을 수 없습니다. '{PPTX_TEMPLATE_PATH}' 파일이 같은 폴더에 있는지 확인해 주세요.")
                return
            
            # --- 1페이지 (Index 0) ---
            remove_unused_textboxes(prs.slides[0], q1_score)
            for shape in prs.slides[0].shapes:
                if shape.has_chart:
                    update_chart_data(shape.chart, averages[0:3])
                    break
                    
            # --- 2페이지 (Index 1) ---
            for shape in prs.slides[1].shapes:
                if shape.has_chart:
                    update_chart_data(shape.chart, averages[3:6])
                    break
                    
            # --- 3페이지 (Index 2) ---
            for shape in prs.slides[2].shapes:
                if shape.has_chart:
                    update_chart_data(shape.chart, averages[6:9])
                    break
                    
            # --- 4페이지 (Index 3) ---
            if len(prs.slides) >= 4:
                remove_unused_textboxes(prs.slides[3], q1_score)
                chart_idx = 0
                for shape in prs.slides[3].shapes:
                    if shape.has_chart:
                        if chart_idx == 0:
                            update_chart_data(shape.chart, averages[0:3])
                        elif chart_idx == 1:
                            update_chart_data(shape.chart, averages[3:6])
                        elif chart_idx == 2:
                            update_chart_data(shape.chart, averages[6:9])
                        chart_idx += 1

            output = io.BytesIO()
            prs.save(output)
            output.seek(0)
            
            st.success("PPT 생성이 완료되었습니다.")
            st.download_button(
                label="📥 PPT 다운로드",
                data=output,
                file_name=f"팀_설문결과_리포트_{selected_date}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )

if __name__ == "__main__":
    main()