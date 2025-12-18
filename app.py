import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from io import BytesIO

st.set_page_config(page_title="Анализатор курсов", layout="wide")
st.title("📊 Отчет по курсам (Регионы / Сертификаты)")

uploaded_file = st.file_uploader("Выберите файл Excel", type=["xlsx", "xls"])

if uploaded_file:
    # Читаем ВЕСЬ файл без ограничений по колонкам, чтобы избежать ошибки Usecols
    df = pd.read_excel(uploaded_file, dtype=str)
    
    # Сразу чистим названия колонок от лишних пробелов
    df.columns = [str(c).strip() for c in df.columns]
    
    # Определяем нужные нам имена
    col_id = "courses.id"
    col_region = "Область"
    col_cert = "Дата получения сертификата"
    col_name = "Наименование_курса"

    # Проверяем, появились ли колонки после чистки
    missing = [c for c in [col_id, col_region, col_cert] if c not in df.columns]
    
    if missing:
        st.error(f"В файле все еще не найдены колонки: {missing}. Проверьте заголовки в Excel.")
    else:
        # Чистим сами данные в колонках
        df[col_region] = df[col_region].str.strip().fillna("Не указано")
        df[col_id] = df[col_id].str.strip()

        mode = st.radio("Параметры фильтрации", ["Все курсы", "По конкретному ID"])
        course_id_input = ""
        if mode == "По конкретному ID":
            course_id_input = st.text_input("Введите courses.id (например: 52)").strip()

        if st.button("📊 Начать анализ"):
            current_course_name = "Все курсы"
            
            if mode == "По конкретному ID" and course_id_input:
                filtered_df = df[df[col_id] == course_id_input]
                if not filtered_df.empty and col_name in filtered_df.columns:
                    current_course_name = filtered_df[col_name].iloc[0]
                title = f"ОТЧЕТ ПО КУРСУ (ID: {course_id_input})"
            else:
                filtered_df = df.copy()
                title = "СВОДНЫЙ ОТЧЕТ ПО ВСЕМ КУРСАМ"

            if filtered_df.empty:
                st.warning("Данные по вашему запросу не найдены.")
            else:
                # Считаем сертификаты: если дата не пустая (не NaN), значит сертификат есть
                filtered_df['has_cert'] = filtered_df[col_cert].notna() & (filtered_df[col_cert] != 'nan')
                
                report = filtered_df.groupby(col_region).agg(
                    total=(col_region, 'count'),
                    with_cert=('has_cert', 'sum')
                ).reset_index()
                
                report['no_cert'] = report['total'] - report['with_cert']
                report = report.sort_values(by='total', ascending=False)
                
                totals = (report['total'].sum(), report['with_cert'].sum(), report['no_cert'].sum())

                st.subheader(title)
                if current_course_name != "Все курсы":
                    st.info(f"Название курса: {current_course_name}")
                
                # Показываем таблицу в браузере
                st.dataframe(report, use_container_width=True)
                
                col1, col2, col3 = st.columns(3)
                col1.metric("Всего человек", totals[0])
                col2.metric("С сертификатом", totals[1])
                col3.metric("Без сертификата", totals[2])

                # Экспорт в Excel
                output = BytesIO()
                wb = Workbook()
                ws = wb.active
                ws.title = "Аналитика"
                ws.append([title])
                ws.append(["ID курса:", course_id_input if course_id_input else "Все"])
                ws.append(["Название курса:", current_course_name])
                ws.append([])
                ws.append(["Область", "Всего людей", "С сертификатом", "Без сертификата"])
                
                for _, row in report.iterrows():
                    ws.append([row['Область'], row['total'], row['with_cert'], row['no_cert']])
                
                ws.append(["ИТОГО", totals[0], totals[1], totals[2]])
                
                for cell in ws[ws.max_row]:
                    cell.font = Font(bold=True)
                
                wb.save(output)
                
                st.download_button(
                    label="💾 Скачать отчет в Excel",
                    data=output.getvalue(),
                    file_name=f"report_{course_id_input if course_id_input else 'all'}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )