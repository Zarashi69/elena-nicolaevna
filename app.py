import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font
from io import BytesIO


st.set_page_config(page_title="Анализатор курсов", layout="wide")

st.title("📊 Отчет по курсам (Регионы / Сертификаты)")

uploaded_file = st.file_uploader("Выберите файл Excel", type=["xlsx", "xls"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]

    col_id = "courses.id"
    col_region = "Область"
    col_cert = "Дата получения сертификата"
    col_name = "Наименование_курса"

    missing = [c for c in [col_id, col_region, col_cert] if c not in df.columns]

    if missing:
        st.error(f"В файле не найдены колонки: {missing}")
    else:
        df[col_region] = df[col_region].str.strip().fillna("Не указано")
        df[col_id] = df[col_id].str.strip()

        mode = st.radio("Параметры фильтрации", ["Все курсы", "По конкретному ID"])

        course_id = ""
        if mode == "По конкретному ID":
            course_id = st.text_input("Введите courses.id").strip()

        if st.button("📊 Начать анализ"):
            current_course_name = "Все курсы"
            filtered_df = df.copy()

            if mode == "По конкретному ID" and course_id:
                filtered_df = df[df[col_id] == course_id]
                if not filtered_df.empty and col_name in filtered_df.columns:
                    current_course_name = filtered_df[col_name].iloc[0]
                title = f"ОТЧЕТ ПО КУРСУ (ID: {course_id})"
            else:
                title = "СВОДНЫЙ ОТЧЕТ ПО ВСЕМ КУРСАМ"

            if filtered_df.empty:
                st.warning("Данные не найдены")
            else:
                filtered_df["has_cert"] = filtered_df[col_cert].notna()
                report = (
                    filtered_df.groupby(col_region)
                    .agg(
                        total=(col_region, "count"),
                        with_cert=("has_cert", "sum"),
                    )
                    .reset_index()
                )

                report["no_cert"] = report["total"] - report["with_cert"]
                report = report.sort_values(by="total", ascending=False)

                totals = (
                    report["total"].sum(),
                    report["with_cert"].sum(),
                    report["no_cert"].sum(),
                )

                st.subheader(title)
                if current_course_name != "Все курсы":
                    st.info(f"Название курса: {current_course_name}")

                st.table(report)

                st.metric("Всего человек", totals[0])
                st.metric("С сертификатом", totals[1])

                output = BytesIO()
                wb = Workbook()
                ws = wb.active
                ws.append([title])
                ws.append(["ID курса:", course_id if course_id else "Все"])
                ws.append(["Название курса:", current_course_name])
                ws.append([])
                ws.append(
                    [
                        "Область",
                        "Всего людей",
                        "С сертификатом",
                        "Без сертификата",
                    ]
                )

                for _, row in report.iterrows():
                    ws.append(
                        [
                            row["Область"],
                            row["total"],
                            row["with_cert"],
                            row["no_cert"],
                        ]
                    )

                ws.append(["ИТОГО", totals[0], totals[1], totals[2]])

                for cell in ws[ws.max_row]:
                    cell.font = Font(bold=True)

                wb.save(output)

                st.download_button(
                    label="💾 Скачать отчет в Excel",
                    data=output.getvalue(),
                    file_name=f"report_{course_id if course_id else 'all'}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )


