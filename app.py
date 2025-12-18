import streamlit as st
import pandas as pd
from openpyxl import Workbook
from io import BytesIO

st.set_page_config(page_title="Анализатор курсов", layout="wide")
st.title("📊 Отчет по курсам (Оптимизировано)")

@st.cache_data(show_spinner=False)
def load_optimized_data(file):
    # Используем движок calamine — он намного быстрее и легче для памяти
    data = pd.read_excel(file, engine='calamine', dtype=str)
    # Убираем пробелы в названиях колонок
    data.columns = [str(c).strip() for c in data.columns]
    return data

uploaded_file = st.file_uploader("Выберите файл Excel", type=["xlsx", "xls"])

if uploaded_file:
    # Используем контейнер, чтобы визуально всё было чисто
    with st.status("🚀 Загрузка и обработка данных...", expanded=True) as status:
        try:
            df = load_optimized_data(uploaded_file)
            status.update(label="✅ Данные готовы!", state="complete", expanded=False)
        except Exception as e:
            st.error(f"Ошибка памяти или формата: {e}")
            st.stop()

    col_id = "courses.id"
    col_region = "Область"
    col_cert = "Дата получения сертификата"
    
    if col_id in df.columns and col_region in df.columns:
        st.divider()
        mode = st.radio("Фильтр:", ["Все", "По ID"], horizontal=True)
        
        c_id = ""
        if mode == "По ID":
            c_id = st.text_input("Введите ID курса").strip()

        if st.button("📊 ПОКАЗАТЬ АНАЛИЗ", type="primary"):
            # Фильтруем
            filtered = df[df[col_id] == c_id].copy() if (mode == "По ID" and c_id) else df.copy()
            
            if filtered.empty:
                st.warning("Ничего не найдено")
            else:
                # Быстрый расчет без лишних колонок
                filtered['has_cert'] = filtered[col_cert].notna()
                res = filtered.groupby(col_region).size().reset_index(name='total')
                certs = filtered[filtered['has_cert']].groupby(col_region).size().reset_index(name='with_cert')
                
                report = pd.merge(res, certs, on=col_region, how='left').fillna(0)
                report['no_cert'] = report['total'] - report['with_cert']
                
                # Итоги
                st.metric("Всего по выборке", int(report['total'].sum()))
                st.dataframe(report.sort_values('total', ascending=False), use_container_width=True)
                
                # Кнопка скачивания появится сразу под таблицей
                output = BytesIO()
                report.to_excel(output, index=False)
                st.download_button("💾 Скачать результат", output.getvalue(), "report.xlsx")
    else:
        st.error(f"Колонки не найдены. Доступны: {list(df.columns[:5])}...")