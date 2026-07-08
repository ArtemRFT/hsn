import io
from datetime import date

import pandas as pd
import streamlit as st

from hsn_report import process, ReportError, PIVOT_COLS

st.set_page_config(
    page_title="ХСН — квартальный отчёт",
    page_icon="🫀",
    layout="centered",
)

st.markdown(
    """
    <style>
    .stApp { background: #EEF2F1; }
    div[data-testid="stMetricValue"] { font-family: 'IBM Plex Mono', monospace; color: #0E6B5C; }
    .block-container { padding-top: 2.5rem; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown("**Инструмент · ХСН**")
st.title("Квартальный отчёт по пациентам с ХСН")
st.write(
    "Загрузите выгрузку «Детализация». Приложение уберёт дубли по СНИЛС и "
    "пациентов с датой смерти, построит сводные таблицы и соберёт готовый "
    "файл со сводками и диаграммами для скачивания."
)

uploaded = st.file_uploader("Файл выгрузки (.xlsx)", type=["xlsx"])

if uploaded is not None:
    title = f"ХСН — отчёт от {date.today().strftime('%d.%m.%Y')}"
    with st.spinner("Обрабатываю файл…"):
        try:
            workbook_bytes, stats, pivots = process(uploaded, title=title)
        except ReportError as e:
            st.error(str(e))
            st.stop()
        except Exception as e:
            st.error(f"Не получилось обработать файл: {e}")
            st.stop()

    st.success(f"Готово. Лист источника: «{stats['sheet_name']}».")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Исходных записей", f"{stats['orig_count']:,}".replace(",", " "))
    c2.metric("Удалено (умершие)", f"{stats['died_count']:,}".replace(",", " "))
    c3.metric("Удалено (дубли СНИЛС)", f"{stats['dup_count']:,}".replace(",", " "))
    c4.metric("Итого пациентов", f"{stats['final_count']:,}".replace(",", " "))

    st.download_button(
        label="⬇️ Скачать готовый xlsx",
        data=workbook_bytes,
        file_name=f"ХСН_обработано_{date.today().isoformat()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )

    st.subheader("Предпросмотр сводных таблиц")
    cols = st.columns(2)
    for i, col in enumerate(PIVOT_COLS):
        vc = pivots[col]
        with cols[i % 2]:
            st.markdown(f"**{col}**")
            if len(vc) == 0:
                st.caption("Нет данных по этому показателю")
            else:
                st.dataframe(vc, hide_index=True, use_container_width=True, height=min(38 * (len(vc) + 1), 260))
else:
    st.info("Ожидаю файл .xlsx с листом «Детализация» и столбцами СНИЛС / Дата смерти.")

st.caption("Обработка выполняется на сервере Streamlit, только пока открыта эта вкладка. Файлы не сохраняются между сессиями.")
