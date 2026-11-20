import io
import zipfile
from datetime import date

import pandas as pd
import streamlit as st

from hsn_report import process, ReportError, PIVOT_COLS, _find_header_row

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
    "Загрузите одну или несколько выгрузок «Детализация». Для каждого файла "
    "приложение уберёт дубли по СНИЛС и пациентов с датой смерти, построит "
    "сводные таблицы и соберёт готовый файл. Файлы обрабатываются независимо: "
    "ошибка в одном не мешает остальным."
)

MO_COL = "МО прикрепления"


def _extract_mo(uploaded, sheet_name: str, header_row: int) -> str:
    """Достаёт первое непустое значение «МО прикрепления» из уже открытого листа."""
    try:
        df = pd.read_excel(uploaded, sheet_name=sheet_name, header=header_row)
    except Exception:
        return ""
    if MO_COL not in df.columns:
        return ""
    vals = df[MO_COL].dropna().astype(str).str.strip()
    vals = vals[vals != ""]
    return vals.iloc[0] if len(vals) else ""


uploaded_files = st.file_uploader(
    "Файлы выгрузок (.xlsx)", type=["xlsx"], accept_multiple_files=True
)

if uploaded_files:
    today_iso = date.today().isoformat()
    results = []  # (orig_name, out_name, bytes, stats, pivots, error, mo)

    progress = st.progress(0.0, text="Обрабатываю файлы…")
    for i, uploaded in enumerate(uploaded_files, start=1):
        orig_name = uploaded.name
        title = f"ХСН — отчёт от {date.today().strftime('%d.%m.%Y')} ({orig_name})"
        try:
            workbook_bytes, stats, pivots = process(uploaded, title=title)
            stem = orig_name.rsplit(".", 1)[0]
            out_name = f"ХСН_обработано_{stem}_{today_iso}.xlsx"

            # uploaded уже прочитан process() — сбрасываем указатель и читаем МО
            uploaded.seek(0)
            mo = _extract_mo(uploaded, stats["sheet_name"], stats["header_row"])
            if not mo:
                mo = stem  # fallback: имя файла без расширения

            results.append((orig_name, out_name, workbook_bytes, stats, pivots, None, mo))
        except ReportError as e:
            results.append((orig_name, None, None, None, None, str(e), ""))
        except Exception as e:
            results.append(
                (orig_name, None, None, None, None, f"Не получилось обработать файл: {e}", "")
            )
        progress.progress(i / len(uploaded_files), text=f"Обработано {i} из {len(uploaded_files)}")
    progress.empty()

    ok = [r for r in results if r[5] is None]
    failed = [r for r in results if r[5] is not None]

    # --- Общая сводка ---
    st.subheader("Итоги по всем файлам")
    total_orig = sum(r[3]["orig_count"] for r in ok)
    total_died = sum(r[3]["died_count"] for r in ok)
    total_dup = sum(r[3]["dup_count"] for r in ok)
    total_final = sum(r[3]["final_count"] for r in ok)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Файлов обработано", f"{len(ok)} / {len(results)}")
    c2.metric("Исходных записей (всего)", f"{total_orig:,}".replace(",", " "))
    c3.metric("Удалено (умершие + дубли)", f"{total_died + total_dup:,}".replace(",", " "))
    c4.metric("Итого пациентов (всего)", f"{total_final:,}".replace(",", " "))

    if failed:
        st.warning(f"Не удалось обработать {len(failed)} файл(ов). Подробности — ниже.")

    # --- ZIP со всеми готовыми файлами ---
    if len(ok) > 1:
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for _, out_name, workbook_bytes, _, _, _, _ in ok:
                zf.writestr(out_name, workbook_bytes.getvalue())
        zip_buf.seek(0)
        st.download_button(
            label="⬇️ Скачать все готовые файлы (ZIP)",
            data=zip_buf,
            file_name=f"ХСН_отчёты_{today_iso}.zip",
            mime="application/zip",
            use_container_width=True,
            type="primary",
            key="download_zip_all",
        )

    # --- По каждому файлу ---
    st.subheader("Отчёты по файлам")
    for idx, (orig_name, out_name, workbook_bytes, stats, pivots, error, mo) in enumerate(results):
        label = f"{mo} — {orig_name}" if mo else orig_name
        with st.expander(label, expanded=(error is not None)):
            if error is not None:
                st.error(error)
                continue

            st.success(f"Готово. Лист источника: «{stats['sheet_name']}». МО: {mo}")

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Исходных записей", f"{stats['orig_count']:,}".replace(",", " "))
            c2.metric("Удалено (умершие)", f"{stats['died_count']:,}".replace(",", " "))
            c3.metric("Удалено (дубли СНИЛС)", f"{stats['dup_count']:,}".replace(",", " "))
            c4.metric("Итого пациентов", f"{stats['final_count']:,}".replace(",", " "))

            st.download_button(
                label=f"⬇️ Скачать {out_name}",
                data=workbook_bytes,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key=f"download_{idx}",
            )

            st.markdown("**Предпросмотр сводных таблиц**")
            cols = st.columns(2)
            for i, col in enumerate(PIVOT_COLS):
                vc = pivots[col]
                with cols[i % 2]:
                    st.markdown(f"**{col}**")
                    if len(vc) == 0:
                        st.caption("Нет данных по этому показателю")
                    else:
                        st.dataframe(
                            vc,
                            hide_index=True,
                            use_container_width=True,
                            height=min(38 * (len(vc) + 1), 260),
                        )

    # --- Итоговая таблица по всем файлам ---
    if ok:
        st.subheader("Итоговая таблица по файлам")

        summary_rows = []
        for num, (orig_name, out_name, workbook_bytes, stats, pivots, error, mo) in enumerate(ok, start=1):
            summary_rows.append({
                "№": num,
                "Наименование организации": mo,
                "Исходно": stats["orig_count"],
                "Кол-во пациентов": stats["final_count"],
            })
        summary_rows.append({
            "№": None,
            "Наименование организации": "ИТОГО",
            "Исходно": total_orig,
            "Кол-во пациентов": total_final,
        })
        summary_df = pd.DataFrame(summary_rows)

        st.dataframe(
            summary_df,
            hide_index=True,
            use_container_width=True,
            column_config={
                "№": st.column_config.NumberColumn("№", width="small", format="%d"),
                "Наименование организации": st.column_config.TextColumn(
                    "Наименование организации", width="large"
                ),
                "Исходно": st.column_config.NumberColumn("Исходно", format="%d"),
                "Кол-во пациентов": st.column_config.NumberColumn(
                    "Кол-во пациентов", format="%d"
                ),
            },
        )
else:
    st.info(
        "Ожидаю один или несколько файлов .xlsx с листом «Детализация» "
        "и столбцами СНИЛС / Дата смерти / МО прикрепления."
    )

st.caption(
    "Обработка выполняется на сервере Streamlit, только пока открыта эта вкладка. "
    "Файлы не сохраняются между сессиями."
)
