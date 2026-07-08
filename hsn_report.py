"""
Логика очистки данных и построения сводных таблиц для отчёта ХСН.
Не зависит от Streamlit — можно использовать и тестировать отдельно.
"""
import io
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference

PIVOT_COLS = [
    'Группа инвалидности',
    'Код федеральной льготы',
    'Код региональной льготы',
    'Стадия ХСН',
    'Функциональный класс ХСН',
    'Вид ХСН по ФВ ЛЖ',
]

SHEET_NAME_MAP = {
    'Группа инвалидности': 'Группа инвалидности',
    'Код федеральной льготы': 'Фед. льгота',
    'Код региональной льготы': 'Рег. льгота',
    'Стадия ХСН': 'Стадия ХСН',
    'Функциональный класс ХСН': 'Функц. класс ХСН',
    'Вид ХСН по ФВ ЛЖ': 'Вид ХСН по ФВ ЛЖ',
}

REQUIRED_COLS = ['СНИЛС', 'Дата смерти']

FONT_NAME = 'Arial'
HEADER_FILL = PatternFill('solid', fgColor='2F5496')
HEADER_FONT = Font(name=FONT_NAME, bold=True, color='FFFFFF', size=11)
TITLE_FONT = Font(name=FONT_NAME, bold=True, size=14, color='2F5496')
SUBTITLE_FONT = Font(name=FONT_NAME, size=10, italic=True, color='595959')
NORMAL_FONT = Font(name=FONT_NAME, size=10)
BOLD_FONT = Font(name=FONT_NAME, size=10, bold=True)
THIN = Side(style='thin', color='BFBFBF')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
ALT_FILL = PatternFill('solid', fgColor='F2F2F2')


class ReportError(Exception):
    pass


def _find_header_row(raw: pd.DataFrame, max_scan_rows: int = 6) -> int:
    """Ищет строку, где встречается заголовок 'СНИЛС', в первых max_scan_rows строках."""
    for i in range(min(max_scan_rows, len(raw))):
        row_values = raw.iloc[i].astype(str).tolist()
        if 'СНИЛС' in row_values:
            return i
    raise ReportError(
        "Не нашёл строку заголовков со столбцом «СНИЛС» в первых "
        f"{max_scan_rows} строках. Проверьте, что загружен правильный файл "
        "с листом «Детализация»."
    )


def load_and_clean(file) -> tuple[pd.DataFrame, dict]:
    """
    file: путь или file-like объект (например, из st.file_uploader).
    Возвращает (очищенный DataFrame, статистика очистки).
    """
    xls = pd.ExcelFile(file)
    sheet_name = next((s for s in xls.sheet_names if 'Детализ' in s), xls.sheet_names[0])

    raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)
    header_row = _find_header_row(raw)

    df = pd.read_excel(xls, sheet_name=sheet_name, header=header_row)
    df = df.dropna(how='all')

    for col in REQUIRED_COLS:
        if col not in df.columns:
            raise ReportError(f"В файле не найден обязательный столбец «{col}».")

    orig_count = len(df)
    died_count = int(df['Дата смерти'].notna().sum())
    df_alive = df[df['Дата смерти'].isna()].copy()

    has_snils = df_alive[df_alive['СНИЛС'].notna()].copy()
    no_snils = df_alive[df_alive['СНИЛС'].isna()].copy()
    dup_count = int(has_snils.duplicated(subset='СНИЛС', keep='first').sum())
    has_snils_dedup = has_snils.drop_duplicates(subset='СНИЛС', keep='first')

    df_clean = pd.concat([has_snils_dedup, no_snils], ignore_index=True)

    stats = {
        'sheet_name': sheet_name,
        'header_row': header_row,
        'orig_count': orig_count,
        'died_count': died_count,
        'dup_count': dup_count,
        'no_snils_count': len(no_snils),
        'final_count': len(df_clean),
    }
    return df_clean, stats


def build_pivots(df_clean: pd.DataFrame) -> dict:
    pivots = {}
    for col in PIVOT_COLS:
        if col not in df_clean.columns:
            pivots[col] = pd.DataFrame(columns=['Значение', 'Кол-во'])
            continue
        vc = df_clean[col].value_counts(dropna=False).reset_index()
        vc.columns = ['Значение', 'Кол-во']
        vc = vc[vc['Значение'].notna()].copy()
        vc['Значение'] = vc['Значение'].astype(str)
        vc = vc.sort_values('Кол-во', ascending=False).reset_index(drop=True)
        pivots[col] = vc
    return pivots


def build_workbook(df_clean: pd.DataFrame, stats: dict, pivots: dict, title: str) -> io.BytesIO:
    wb = Workbook()

    # --- Сводка ---
    ws = wb.active
    ws.title = 'Сводка'
    ws.sheet_view.showGridLines = False
    ws['B2'] = title
    ws['B2'].font = TITLE_FONT
    ws['B3'] = 'Отчёт после очистки данных (дубли по СНИЛС и умершие пациенты удалены)'
    ws['B3'].font = SUBTITLE_FONT

    info_rows = [
        ('Исходное количество записей', stats['orig_count']),
        ('Удалено (пациенты с датой смерти)', stats['died_count']),
        ('Удалено (дубли по СНИЛС)', stats['dup_count']),
        ('Записей без СНИЛС (оставлены как уникальные)', stats['no_snils_count']),
        ('Итоговое количество уникальных пациентов', stats['final_count']),
    ]
    r = 5
    ws[f'B{r}'] = 'Показатель'; ws[f'C{r}'] = 'Значение'
    ws[f'B{r}'].font = HEADER_FONT; ws[f'C{r}'].font = HEADER_FONT
    ws[f'B{r}'].fill = HEADER_FILL; ws[f'C{r}'].fill = HEADER_FILL
    r += 1
    for label, val in info_rows:
        ws[f'B{r}'] = label
        ws[f'C{r}'] = val
        ws[f'B{r}'].font = NORMAL_FONT
        ws[f'C{r}'].font = BOLD_FONT
        ws[f'B{r}'].border = BORDER
        ws[f'C{r}'].border = BORDER
        ws[f'C{r}'].alignment = Alignment(horizontal='center')
        r += 1

    ws.column_dimensions['A'].width = 3
    ws.column_dimensions['B'].width = 50
    ws.column_dimensions['C'].width = 15

    r += 2
    ws[f'B{r}'] = 'Содержание отчёта:'
    ws[f'B{r}'].font = BOLD_FONT
    r += 1
    sheet_desc = [('Детализация (очищено)', 'Полный очищенный список пациентов')]
    sheet_desc += [(SHEET_NAME_MAP[c], f'Сводная таблица: {c}') for c in PIVOT_COLS]
    for name, desc in sheet_desc:
        ws[f'B{r}'] = f'• {name} — {desc}'
        ws[f'B{r}'].font = NORMAL_FONT
        r += 1

    # --- Детализация ---
    ws2 = wb.create_sheet('Детализация (очищено)')
    for j, col_name in enumerate(df_clean.columns, start=1):
        c = ws2.cell(row=1, column=j, value=col_name)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border = BORDER
    ws2.row_dimensions[1].height = 30

    for i, row in enumerate(df_clean.itertuples(index=False), start=2):
        for j, val in enumerate(row, start=1):
            if pd.isna(val):
                val = None
            c = ws2.cell(row=i, column=j, value=val)
            c.font = NORMAL_FONT
            c.border = BORDER
            if i % 2 == 0:
                c.fill = ALT_FILL

    for j, col_name in enumerate(df_clean.columns, start=1):
        lengths = [len(str(col_name))]
        for v in df_clean.iloc[:, j - 1]:
            if pd.notna(v):
                lengths.append(len(str(v)))
        max_len = max(lengths) if lengths else 10
        ws2.column_dimensions[get_column_letter(j)].width = min(max(max_len + 2, 12), 40)

    ws2.freeze_panes = 'A2'

    # --- Сводные таблицы ---
    for col in PIVOT_COLS:
        vc = pivots[col]
        sname = SHEET_NAME_MAP[col]
        ws_p = wb.create_sheet(sname[:31])
        ws_p.sheet_view.showGridLines = False
        ws_p['B2'] = col
        ws_p['B2'].font = TITLE_FONT
        ws_p['B3'] = f"Всего уникальных пациентов: {stats['final_count']}"
        ws_p['B3'].font = SUBTITLE_FONT

        start_row = 5

        if len(vc) == 0:
            ws_p.cell(row=start_row, column=2, value='Нет данных по этому показателю').font = NORMAL_FONT
            ws_p.column_dimensions['A'].width = 3
            ws_p.column_dimensions['B'].width = 55
            ws_p.column_dimensions['C'].width = 12
            continue

        ws_p.cell(row=start_row, column=2, value='Значение').font = HEADER_FONT
        ws_p.cell(row=start_row, column=3, value='Кол-во').font = HEADER_FONT
        ws_p.cell(row=start_row, column=2).fill = HEADER_FILL
        ws_p.cell(row=start_row, column=3).fill = HEADER_FILL
        ws_p.cell(row=start_row, column=2).alignment = Alignment(horizontal='left')
        ws_p.cell(row=start_row, column=3).alignment = Alignment(horizontal='center')

        for idx, rowdata in vc.iterrows():
            rr = start_row + 1 + idx
            cval = ws_p.cell(row=rr, column=2, value=rowdata['Значение'])
            ccnt = ws_p.cell(row=rr, column=3, value=int(rowdata['Кол-во']))
            cval.font = NORMAL_FONT
            ccnt.font = NORMAL_FONT
            cval.border = BORDER
            ccnt.border = BORDER
            ccnt.alignment = Alignment(horizontal='center')
            if idx % 2 == 1:
                cval.fill = ALT_FILL
                ccnt.fill = ALT_FILL

        total_row = start_row + 1 + len(vc)
        ws_p.cell(row=total_row, column=2, value='Итого').font = BOLD_FONT
        tot_cell = ws_p.cell(row=total_row, column=3, value=f'=SUM(C{start_row + 1}:C{total_row - 1})')
        tot_cell.font = BOLD_FONT
        ws_p.cell(row=total_row, column=2).border = BORDER
        tot_cell.border = BORDER
        tot_cell.alignment = Alignment(horizontal='center')

        ws_p.column_dimensions['A'].width = 3
        ws_p.column_dimensions['B'].width = 55
        ws_p.column_dimensions['C'].width = 12

        chart = BarChart()
        chart.title = col
        chart.x_axis.delete = False
        chart.y_axis.delete = False
        chart.style = 10
        data = Reference(ws_p, min_col=3, min_row=start_row, max_row=total_row - 1)
        cats = Reference(ws_p, min_col=2, min_row=start_row + 1, max_row=total_row - 1)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        chart.legend = None
        chart.height = 8
        chart.width = 18
        ws_p.add_chart(chart, f'E{start_row}')

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def process(file, title: str = 'ХСН — квартальный отчёт'):
    """Единая точка входа: загрузка → очистка → сводные → итоговый xlsx (BytesIO)."""
    df_clean, stats = load_and_clean(file)
    pivots = build_pivots(df_clean)
    workbook_bytes = build_workbook(df_clean, stats, pivots, title)
    return workbook_bytes, stats, pivots
