from __future__ import annotations

import argparse
import csv
import io
import json
import re
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ALIASES = {
    "date": {"дата", "date", "дата продажи"},
    "category": {"категория", "category", "категория товара", "группа"},
    "manager": {"менеджер", "manager", "сотрудник"},
    "qty": {"кол-во", "количество", "qty", "quantity", "кол-во, шт"},
    "amount": {"сумма", "сумма, руб", "сумма, ₽", "сумма руб", "sum", "amount", "выручка"},
}

REQUIRED = ("date", "category", "amount")

HEADER_FILL = PatternFill("solid", fgColor="217346")
HEADER_FONT = Font(bold=True, color="FFFFFF")
TOTAL_FILL = PatternFill("solid", fgColor="F1F8F4")
THIN = Side(style="thin", color="E4E9F0")
BORDER = Border(bottom=THIN)
MONEY = "# ##0.00"


def decode(raw: bytes) -> str:
    for encoding in ("utf-8-sig", "utf-8", "cp1251"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw.decode("cp1251", errors="replace")


def sniff_delimiter(text: str) -> str:
    sample = "\n".join(text.splitlines()[:10])
    try:
        return csv.Sniffer().sniff(sample, delimiters=";,\t").delimiter
    except csv.Error:
        return ";" if sample.count(";") >= sample.count(",") else ","


def read_csv(path: Path) -> pd.DataFrame:
    text = decode(path.read_bytes())
    return pd.read_csv(io.StringIO(text), sep=sniff_delimiter(text), dtype=str)


def read_excel_any(path: Path) -> pd.DataFrame:
    probe = pd.read_excel(path, header=None, dtype=str)
    header_row = None
    for index, row in probe.iterrows():
        cells = {str(value).strip().lower().replace("ё", "е") for value in row.dropna()}
        if any("дата" in cell or "date" in cell for cell in cells) and any(
            "категор" in cell or "сумм" in cell for cell in cells
        ):
            header_row = index
            break
    if header_row is None:
        header_row = 0
    return pd.read_excel(path, header=header_row, dtype=str)


def normalize_columns(frame: pd.DataFrame) -> pd.DataFrame:
    mapping = {}
    for column in frame.columns:
        key = re.sub(r"\s+", " ", str(column).strip().lower().replace("ё", "е"))
        for canonical, aliases in ALIASES.items():
            if key in aliases:
                mapping[column] = canonical
                break
    frame = frame.rename(columns=mapping)
    missing = [name for name in REQUIRED if name not in frame.columns]
    if missing:
        raise ValueError(f"нет обязательных колонок: {', '.join(missing)}")
    for optional in ("manager", "qty"):
        if optional not in frame.columns:
            frame[optional] = None
    return frame


def parse_money(value) -> float | None:
    if value is None:
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    text = str(value).replace("\xa0", " ").replace(" ", "")
    text = text.replace("₽", "").replace("руб.", "").replace("руб", "")
    text = text.replace(",", ".")
    text = re.sub(r"[^0-9.\-]", "", text)
    if text in ("", "-", ".", "-."):
        return None
    try:
        return float(text)
    except ValueError:
        return None


FORMATS = ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y", "%d.%m.%y", "%Y.%m.%d")


def parse_dates(series: pd.Series) -> pd.Series:
    text = series.astype(str).str.strip()
    result = pd.Series(pd.NaT, index=series.index, dtype="datetime64[ns]")
    for fmt in FORMATS:
        result = result.fillna(pd.to_datetime(text, format=fmt, errors="coerce"))
    return result


def clean_frame(frame: pd.DataFrame, source: str) -> pd.DataFrame:
    frame = frame.copy()
    frame["category"] = (
        frame["category"].astype(str).str.replace("\xa0", " ", regex=False)
        .str.strip().str.replace(r"\s+", " ", regex=True).str.title()
    )
    frame["date"] = parse_dates(frame["date"])
    frame["amount"] = frame["amount"].map(parse_money)
    frame["qty"] = pd.to_numeric(frame["qty"], errors="coerce").fillna(0).astype(int)
    frame["manager"] = frame["manager"].fillna("—").astype(str).str.strip()

    before = len(frame)
    frame = frame.dropna(subset=["date", "amount"])
    dropped_total_rows = before - len(frame)
    frame = frame[frame["amount"] > 0]
    frame = frame.drop_duplicates(subset=["date", "category", "manager", "qty", "amount"])
    frame["source"] = source
    frame.attrs["dropped_total_rows"] = dropped_total_rows
    return frame.sort_values("date").reset_index(drop=True)


def collect(input_dir: Path) -> pd.DataFrame:
    files = sorted([*input_dir.glob("*.csv"), *input_dir.glob("*.xlsx")])
    if not files:
        raise SystemExit(f"в {input_dir} не найдено csv или xlsx")
    parts = []
    for path in files:
        raw = read_csv(path) if path.suffix.lower() == ".csv" else read_excel_any(path)
        frame = clean_frame(normalize_columns(raw), path.name)
        print(f"{path.name}: строк {len(frame)}, отброшено пустых/итоговых "
              f"{frame.attrs.get('dropped_total_rows', 0)}")
        parts.append(frame)
    return pd.concat(parts, ignore_index=True)


def build_pivot(frame: pd.DataFrame) -> pd.DataFrame:
    frame = frame.copy()
    frame["month"] = frame["date"].dt.to_period("M").astype(str)
    pivot = frame.pivot_table(index="month", columns="category",
                              values="amount", aggfunc="sum", fill_value=0)
    pivot["Итого"] = pivot.sum(axis=1)
    pivot.loc["Итого"] = pivot.sum()
    return pivot.round(2)


def write_report(path: Path, frame: pd.DataFrame, pivot: pd.DataFrame) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Сводная"
    sheet["A1"] = "Отчёт по продажам"
    sheet["A1"].font = Font(bold=True, size=14)

    sheet["A3"] = "Месяц"
    for offset, category in enumerate(pivot.columns, start=2):
        sheet.cell(row=3, column=offset, value=category)
    for column in range(1, len(pivot.columns) + 2):
        cell = sheet.cell(row=3, column=column)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center")

    for row_offset, (month, values) in enumerate(pivot.iterrows(), start=4):
        sheet.cell(row=row_offset, column=1, value=str(month))
        for column_offset, value in enumerate(values, start=2):
            cell = sheet.cell(row=row_offset, column=column_offset, value=float(value))
            cell.number_format = MONEY
        if str(month) == "Итого" or month == pivot.index[-1]:
            for column in range(1, len(pivot.columns) + 2):
                cell = sheet.cell(row=row_offset, column=column)
                cell.font = Font(bold=True)
                cell.fill = TOTAL_FILL

    last_row = 3 + len(pivot)
    sheet.freeze_panes = "A4"
    sheet.column_dimensions["A"].width = 14
    for column in range(2, len(pivot.columns) + 2):
        sheet.column_dimensions[get_column_letter(column)].width = 16
        for row in range(4, last_row + 1):
            sheet.cell(row=row, column=column).border = BORDER

    chart = BarChart()
    chart.type = "col"
    chart.title = "Выручка по месяцам"
    chart.y_axis.numFmt = MONEY
    data = Reference(sheet, min_col=2, max_col=len(pivot.columns) - 1,
                     min_row=3, max_row=last_row - 1)
    categories = Reference(sheet, min_col=1, min_row=4, max_row=last_row - 1)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)
    chart.height = 8
    chart.width = 18
    sheet.add_chart(chart, f"A{last_row + 2}")

    data_sheet = workbook.create_sheet("Данные")
    columns = ["date", "category", "manager", "qty", "amount", "source"]
    titles = ["Дата", "Категория", "Менеджер", "Кол-во", "Сумма", "Источник"]
    for index, title in enumerate(titles, start=1):
        cell = data_sheet.cell(row=1, column=index, value=title)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
    for row_index, row in enumerate(frame[columns].itertuples(index=False), start=2):
        data_sheet.cell(row=row_index, column=1, value=row.date.to_pydatetime())
        data_sheet.cell(row=row_index, column=1).number_format = "DD.MM.YYYY"
        data_sheet.cell(row=row_index, column=2, value=row.category)
        data_sheet.cell(row=row_index, column=3, value=row.manager)
        data_sheet.cell(row=row_index, column=4, value=row.qty)
        cell = data_sheet.cell(row=row_index, column=5, value=float(row.amount))
        cell.number_format = MONEY
        data_sheet.cell(row=row_index, column=6, value=row.source)
    data_sheet.auto_filter.ref = f"A1:F{len(frame) + 1}"
    data_sheet.freeze_panes = "A2"
    for index in range(1, 7):
        data_sheet.column_dimensions[get_column_letter(index)].width = 18

    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)


def main() -> None:
    parser = argparse.ArgumentParser(description="Собирает сводный отчёт из выгрузок Excel и CSV")
    parser.add_argument("--input", type=Path, default=Path("data/raw"))
    parser.add_argument("--output", type=Path, default=Path("outputs/report.xlsx"))
    args = parser.parse_args()

    frame = collect(args.input)
    pivot = build_pivot(frame)
    write_report(args.output, frame, pivot)

    summary = {
        "files": sorted({name for name in frame["source"]}),
        "rows": int(len(frame)),
        "total": round(float(frame["amount"].sum()), 2),
        "by_month": {str(k): round(float(v), 2) for k, v in pivot["Итого"].iloc[:-1].items()},
    }
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    print(f"готово: {args.output}")


if __name__ == "__main__":
    main()
