import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Анализатор", layout="centered")
st.title("📊 Отчет по курсам (Оптимизированный)")

# Функция загрузки с использованием Calamine (самый быстрый движок)
@st.cache_data(show_spinner=False)
def load_data(file):
    try:
        # Читаем только нужные колонки, чтобы сэкономить 80% памяти
        target_cols = ["courses.id", "Область", "Дата получения сертификата", "Наименование_курса"]
        df = pd.read_excel(file, engine='calamine', dtype=str)
        
        # Чистим названия колонок
        df.columns = [str(c).strip() for c in df.columns]
        
        # Оставляем только те колонки, которые реально нужны для анализа
        existing_cols = [c for c in target_cols if c in df.columns]
        return df[existing_cols]
    except Exception as e:
        return str(e)

uploaded_file = st.file_uploader("Загрузите файл", type=["xlsx"])

if uploaded_file:
    with st.spinner('⏳ Обработка данных... Пожалуйста, не закрывайте вкладку'):
        data = load_data(uploaded_file)
    
    if isinstance(data, str):
        st.error(f"Ошибка: {data}")
    else:
        st.success("✅ Данные загружены!")
        
        mode = st.radio("Режим:", ["Все курсы", "По ID"], horizontal=True)
        c_id = st.text_input("Введите ID курса") if mode == "По ID" else None

        if st.button("📊 Сгенерировать отчет"):
            # Фильтрация
            filtered = data[data["courses.id"] == c_id.strip()] if c_id else data
            
            if filtered.empty:
                st.warning("Ничего не найдено")
            else:
                # Считаем итоги
                filtered['has_cert'] = filtered["Дата получения сертификата"].notna()
                
                report = filtered.groupby("Область").agg(
                    total=("Область", "count"),
                    with_cert=("has_cert", "sum")
                ).reset_index()
                
                report["no_cert"] = report["total"] - report["with_cert"]
                report = report.sort_values("total", ascending=False)

                # Вывод
                st.write(f"### Итоги по выбору:")
                st.metric("Всего студентов", len(filtered))
                st.dataframe(report, use_container_width=True)

                # Простая кнопка скачивания
                output = BytesIO()
                report.to_excel(output, index=False)
                st.download_button("💾 Скачать Excel", output.getvalue(), "report.xlsx")