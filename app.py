import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Анализатор (Fast Mode)", layout="centered")
st.title("🚀 Быстрый отчет (Предфильтрация)")

st.info("Этот режим экономит память: мы ищем только нужный ID сразу при чтении.")

# Поле ввода ID ДО загрузки или вместе с ней
course_id = st.text_input("1. Введите ID курса (например, 52)", key="course_id").strip()

uploaded_file = st.file_uploader("2. Загрузите файл Excel", type=["xlsx"])

if uploaded_file and course_id:
    with st.status("🔍 Ищу данные по курсу " + course_id + "...") as status:
        try:
            # Используем calamine для скорости
            # Читаем только нужные колонки
            df_iter = pd.read_excel(uploaded_file, engine='calamine', dtype=str)
            
            # Чистим заголовки
            df_iter.columns = [str(c).strip() for c in df_iter.columns]
            
            # Сразу фильтруем, чтобы не хранить лишнее в памяти
            col_id = "courses.id"
            if col_id in df_iter.columns:
                filtered_df = df_iter[df_iter[col_id] == course_id].copy()
                
                if filtered_df.empty:
                    st.warning(f"Курс с ID {course_id} не найден в файле.")
                    st.stop()
                
                status.update(label="✅ Данные найдены!", state="complete")
            else:
                st.error(f"Колонка {col_id} не найдена!")
                st.stop()
                
        except Exception as e:
            st.error(f"Ошибка: {e}")
            st.stop()

    # Анализ только отфильтрованных данных
    col_region = "Область"
    col_cert = "Дата получения сертификата"
    
    if col_region in filtered_df.columns:
        # Логика сертификатов
        filtered_df['has_cert'] = filtered_df[col_cert].notna() & (filtered_df[col_cert].astype(str).str.lower() != 'nan')
        
        report = filtered_df.groupby(col_region).agg(
            total=(col_region, 'count'),
            with_cert=('has_cert', 'sum')
        ).reset_index()
        
        report['no_cert'] = report['total'] - report['with_cert']
        report = report.sort_values('total', ascending=False)

        # Интерфейс
        st.divider()
        st.subheader(f"Результаты для курса №{course_id}")
        
        m1, m2, m3 = st.columns(3)
        m1.metric("Всего чел.", int(report['total'].sum()))
        m2.metric("С сертификатом", int(report['with_cert'].sum()))
        m3.metric("Без сертификата", int(report['no_cert'].sum()))

        st.dataframe(report, use_container_width=True)

        # Скачивание
        output = BytesIO()
        report.to_excel(output, index=False)
        st.download_button("💾 Скачать отчет по ID " + course_id, output.getvalue(), f"report_{course_id}.xlsx")
    else:
        st.error("Колонка 'Область' не найдена.")

elif uploaded_file and not course_id:
    st.warning("⚠️ Сначала введите ID курса в поле выше, чтобы начать поиск.")