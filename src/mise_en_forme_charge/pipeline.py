"""Pipeline configurable pour reorganiser les commandes en cours."""

from __future__ import annotations

import csv
import json
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter


@dataclass
class TransformResult:
    """Resultat du traitement V1."""

    rows_total: int
    rows_kept: int
    output_path: Path
    output_sheet: str


@dataclass
class Profile:
    """Definition d'un profil de transformation JSON."""

    name: str
    expected_columns: list[str]
    computed_columns: dict[str, dict[str, Any]]
    filters: list[dict[str, Any]]
    output_sheet: str
    column_order: list[str]
    drop_columns: set[str]
    column_format: dict[str, dict[str, Any]]
    create_excel_table: bool
    sort_by_week: bool
    create_weekly_summary: bool
    summary_sheet_name: str


def _first_word(value: str) -> str:
    value = value.strip()
    if not value:
        return ""
    return value.split()[0]


def _parse_french_date(value: Any) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def _parse_profile(path: Path) -> Profile:
    with path.open("r", encoding="utf-8") as handle:
        payload = json.load(handle)

    output = payload.get("output", {})
    return Profile(
        name=str(payload.get("profile", path.stem)),
        expected_columns=list(payload.get("expected_columns", [])),
        computed_columns=dict(payload.get("computed_columns", {})),
        filters=list(payload.get("filters", [])),
        output_sheet=str(output.get("sheet_name", "Resultat")),
        column_order=list(output.get("column_order", [])),
        drop_columns=set(output.get("drop_columns", [])),
        column_format=dict(output.get("column_format", {})),
        create_excel_table=bool(output.get("create_excel_table", True)),
        sort_by_week=bool(output.get("sort_by_week", True)),
        create_weekly_summary=bool(output.get("create_weekly_summary", True)),
        summary_sheet_name=str(output.get("summary_sheet_name", "CHARGE_HEBDO")),
    )


def _normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _validate_structure(headers: list[str], expected_columns: list[str]) -> None:
    missing = [col for col in expected_columns if col not in headers]
    if missing:
        listed = ", ".join(missing)
        raise ValueError(f"Structure invalide: colonnes manquantes: {listed}")


def _read_rows(path: Path) -> tuple[list[str], list[dict[str, Any]]]:
    suffix = path.suffix.lower()
    if suffix == ".txt":
        with path.open("r", encoding="cp1252", errors="replace", newline="") as handle:
            reader = csv.DictReader(handle, delimiter="\t")
            headers = [h.strip() for h in (reader.fieldnames or [])]
            rows = []
            for raw_row in reader:
                row: dict[str, Any] = {}
                for key, value in raw_row.items():
                    if key is not None:
                        row[key.strip()] = value
                rows.append(row)
            return headers, rows

    if suffix not in {".xlsx", ".xlsm"}:
        raise ValueError(f"Format non supporte: {path.suffix}")

    workbook = load_workbook(path, read_only=True, data_only=True, keep_vba=(suffix == ".xlsm"))
    sheet = workbook.active
    values = sheet.iter_rows(values_only=True)
    first_row = next(values, None)
    if first_row is None:
        return [], []
    headers = [str(cell).strip() if cell is not None else "" for cell in first_row]
    rows = []
    for line in values:
        row = {headers[idx]: line[idx] for idx in range(len(headers)) if headers[idx]}
        rows.append(row)
    return headers, rows


def _match_filter(value: Any, rule: dict[str, Any]) -> bool:
    text = _normalize_text(value)
    op = str(rule.get("operator", "")).lower()
    target = _normalize_text(rule.get("value"))

    if op == "equals":
        return text.upper() == target.upper()
    if op == "not_equals":
        if rule.get("allow_empty") and text == "":
            return True
        return text.upper() != target.upper()
    if op == "not_empty":
        return text != ""

    raise ValueError(f"Operateur de filtre non supporte: {op}")


def _is_row_kept(row: dict[str, Any], rules: list[dict[str, Any]]) -> bool:
    for rule in rules:
        col = str(rule.get("column", ""))
        if col == "":
            continue
        if not _match_filter(row.get(col), rule):
            return False
    return True


def _compute_value(row: dict[str, Any], rule: dict[str, Any]) -> Any:
    source = str(rule.get("source", ""))
    transform = str(rule.get("transform", ""))
    raw = row.get(source)

    if transform == "iso_week":
        dt = _parse_french_date(raw)
        return dt.isocalendar().week if dt else ""
    if transform == "first_word":
        return _first_word(_normalize_text(raw))

    raise ValueError(f"Transformation non supportee: {transform}")


def _transform_row(row: dict[str, Any], computed_columns: dict[str, dict[str, Any]]) -> dict[str, Any]:
    transformed = dict(row)
    for new_col, rule in computed_columns.items():
        transformed[new_col] = _compute_value(row, rule)
    return transformed


def _ordered_columns(
    headers: list[str],
    transformed_rows: list[dict[str, Any]],
    preferred_order: list[str],
    drop_columns: set[str],
) -> list[str]:
    seen = set()
    ordered = []

    for col in preferred_order:
        if any(col in row for row in transformed_rows):
            ordered.append(col)
            seen.add(col)

    for col in headers:
        if col and col not in seen and col not in drop_columns:
            ordered.append(col)
            seen.add(col)

    for row in transformed_rows:
        for col in row.keys():
            if col not in seen and col not in drop_columns:
                ordered.append(col)
                seen.add(col)

    return ordered


def _format_cell_value(column: str, value: Any, column_format: dict[str, dict[str, Any]]) -> Any:
    fmt = column_format.get(column, {})
    if fmt.get("date_without_time"):
        dt = _parse_french_date(value)
        if dt:
            return dt.strftime("%d/%m/%Y")
    return value


def _parse_quantity(value: Any) -> float:
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip().replace(" ", "")
    if text == "":
        return 0.0
    text = text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return 0.0


def _type_sort_key(type_name: str) -> tuple[int, str]:
    upper = type_name.upper()
    if upper.startswith("MANCHE"):
        return (0, upper)
    if upper.startswith("POCHE"):
        return (1, upper)
    return (2, upper)


def _week_key(row: dict[str, Any]) -> tuple[int, int] | None:
    dt_liv = _parse_french_date(row.get("DT_LIV_CONFIRMEE"))
    if dt_liv is None:
        return None
    iso = dt_liv.isocalendar()
    return (iso.year, iso.week)


def _week_label(week: tuple[int, int]) -> str:
    _, iso_week = week
    return f"S{iso_week:02d}"


def _friday_label(week: tuple[int, int]) -> str:
    year, iso_week = week
    friday = date.fromisocalendar(year, iso_week, 5)
    return friday.strftime("%d/%m")


def _build_weekly_matrix(
    rows: list[dict[str, Any]],
    weeks: list[tuple[int, int]],
) -> tuple[list[str], dict[str, dict[tuple[int, int], float]]]:
    matrix: dict[str, dict[tuple[int, int], float]] = defaultdict(lambda: defaultdict(float))
    types_set: set[str] = set()

    for row in rows:
        week = _week_key(row)
        if week is None:
            continue
        type_name = _normalize_text(row.get("TYPE")).upper()
        if type_name == "":
            type_name = "<SANS_TYPE>"
        qty = _parse_quantity(row.get("RESTE_A_LIV_UV"))
        matrix[type_name][week] += qty
        types_set.add(type_name)

    ordered_types = sorted(types_set, key=_type_sort_key)

    # Ensure every week key exists so row rendering is straightforward.
    for type_name in ordered_types:
        for week in weeks:
            _ = matrix[type_name][week]

    return ordered_types, matrix


def _write_weekly_section(
    sheet: Any,
    start_row: int,
    rows: list[dict[str, Any]],
    weeks: list[tuple[int, int]],
    table_name: str,
) -> int:
    current_iso = date.today().isocalendar()
    current_week_key = (current_iso.year, current_iso.week)
    yellow_pale = PatternFill(fill_type="solid", fgColor="FFFFF8DC")
    green_pale = PatternFill(fill_type="solid", fgColor="FFF0FFF0")
    black_border = Border(
        left=Side(style="thin", color="FF000000"),
        right=Side(style="thin", color="FF000000"),
        top=Side(style="thin", color="FF000000"),
        bottom=Side(style="thin", color="FF000000"),
    )

    friday_row = start_row
    table_header_row = start_row + 1
    data_start_row = start_row + 2

    sheet.cell(friday_row, 1, "Date vendredi")
    sheet.cell(table_header_row, 1, "TYPE")

    for idx, week in enumerate(weeks, start=2):
        sheet.cell(friday_row, idx, _friday_label(week))
        sheet.cell(table_header_row, idx, _week_label(week))

    ordered_types, matrix = _build_weekly_matrix(rows, weeks)

    current_row = data_start_row
    for type_name in ordered_types:
        sheet.cell(current_row, 1, type_name)
        emphasized = type_name.upper().startswith("MANCHE") or type_name.upper().startswith("POCHE")

        if emphasized:
            sheet.cell(current_row, 1).font = Font(bold=True, size=12)
            sheet.row_dimensions[current_row].height = 20

        for idx, week in enumerate(weeks, start=2):
            value = matrix[type_name][week]
            if abs(value) < 1e-9:
                sheet.cell(current_row, idx, "")
            else:
                cell_value = int(round(value))
                sheet.cell(current_row, idx, cell_value)
                sheet.cell(current_row, idx).number_format = "0"
            
            if emphasized:
                sheet.cell(current_row, idx).font = Font(bold=True, size=12)
            else:
                sheet.cell(current_row, idx).font = Font(bold=True, size=10)
        current_row += 1

    end_row = max(current_row - 1, table_header_row)
    end_col = 1 + len(weeks)

    # Ne pas créer de table pour éviter la corruption Excel
    # Les données sont affichées directement sans format tableau

    sheet.column_dimensions["A"].width = 23
    for idx in range(2, end_col + 1):
        letter = get_column_letter(idx)
        sheet.column_dimensions[letter].width = 8

    for row_idx in range(friday_row, end_row + 1):
        sheet.cell(row_idx, 1).alignment = Alignment(horizontal="left", vertical="center")
        for col_idx in range(2, end_col + 1):
            sheet.cell(row_idx, col_idx).alignment = Alignment(horizontal="center", vertical="center")

    # Appliquer les bordures noires à TOUTES les cellules
    for row_idx in range(friday_row, end_row + 1):
        for col_idx in range(1, end_col + 1):
            sheet.cell(row_idx, col_idx).border = black_border

    for idx, week in enumerate(weeks, start=2):
        if week < current_week_key:
            fill = yellow_pale
        elif week == current_week_key:
            fill = green_pale
        else:
            fill = None

        if fill is None:
            continue
        for row_idx in range(data_start_row, end_row + 1):
            sheet.cell(row_idx, idx).fill = fill

    for col_idx in range(1, end_col + 1):
        sheet.cell(friday_row, col_idx).font = Font(bold=False, size=10)
        header_cell = sheet.cell(table_header_row, col_idx)
        header_cell.font = Font(bold=True, size=11, color="FFFFFFFF")
        header_cell.fill = PatternFill(fill_type="solid", fgColor="FF404040")
    
    sheet.row_dimensions[table_header_row].height = 20

    return end_row + 2


def _write_weekly_summary_sheet(
    workbook: Any,
    sheet_name: str,
    transformed_rows: list[dict[str, Any]],
) -> None:
    if sheet_name in workbook.sheetnames:
        del workbook[sheet_name]

    sheet = workbook.create_sheet(sheet_name)
    weeks = sorted({wk for row in transformed_rows if (wk := _week_key(row)) is not None})

    is_willmark = lambda row: "WILLMARK" in _normalize_text(row.get("NOMREDCLI_CDE")).upper()
    rows_main = [row for row in transformed_rows if not is_willmark(row)]
    rows_willmark = [row for row in transformed_rows if is_willmark(row)]

    _write_weekly_section(
        sheet,
        start_row=1,
        rows=rows_main,
        weeks=weeks,
        table_name="TableChargeMain",
    )

    if rows_willmark:
        stock_sheet_name = "CMD_STOCK"
        if stock_sheet_name in workbook.sheetnames:
            del workbook[stock_sheet_name]
        
        stock_sheet = workbook.create_sheet(stock_sheet_name)
        _write_weekly_section(
            stock_sheet,
            start_row=1,
            rows=rows_willmark,
            weeks=weeks,
            table_name="TableChargeWillmark",
        )

def _write_retard_sheet(
    workbook: Any,
    sheet_name: str,
    transformed_rows: list[dict[str, Any]],
    columns: list[str],
    column_format: dict[str, dict[str, Any]],
) -> None:
    """Cree la feuille Retard avec synthese hebdo et details OF."""
    current_iso = date.today().isocalendar()
    current_week_key = (current_iso.year, current_iso.week)

    if sheet_name in workbook.sheetnames:
        del workbook[sheet_name]

    sheet = workbook.create_sheet(sheet_name)

    retard_rows: list[dict[str, Any]] = []
    for row in transformed_rows:
        week_key = _week_key(row)
        if week_key is not None and week_key < current_week_key:
            retard_rows.append(row)

    black_border = Border(
        left=Side(style="thin", color="FF000000"),
        right=Side(style="thin", color="FF000000"),
        top=Side(style="thin", color="FF000000"),
        bottom=Side(style="thin", color="FF000000"),
    )
    header_fill = PatternFill(fill_type="solid", fgColor="FF404040")

    # Bloc gauche: S / M / P / A par semaine + ligne TT
    for col_idx, label in enumerate(["S", "M", "P", "A"], start=1):
        cell = sheet.cell(1, col_idx, label)
        cell.font = Font(bold=True, size=11, color="FFFFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = black_border
        sheet.column_dimensions[get_column_letter(col_idx)].width = 8

    sheet.row_dimensions[1].height = 20

    week_matrix: dict[tuple[int, int], dict[str, float]] = defaultdict(
        lambda: {"M": 0.0, "P": 0.0, "A": 0.0}
    )
    for row in retard_rows:
        week_key = _week_key(row)
        if week_key is None:
            continue
        qty = _parse_quantity(row.get("RESTE_A_LIV_UV"))
        type_name = _normalize_text(row.get("TYPE")).upper()
        if type_name.startswith("MANCHE"):
            week_matrix[week_key]["M"] += qty
        elif type_name.startswith("POCHE"):
            week_matrix[week_key]["P"] += qty
        else:
            week_matrix[week_key]["A"] += qty

    ordered_weeks = sorted(week_matrix.keys())
    row_idx = 2
    total_m = 0.0
    total_p = 0.0
    total_a = 0.0

    for week in ordered_weeks:
        m_value = week_matrix[week]["M"]
        p_value = week_matrix[week]["P"]
        a_value = week_matrix[week]["A"]
        total_m += m_value
        total_p += p_value
        total_a += a_value

        values = [
            _week_label(week),
            int(round(m_value)) if m_value > 0 else "",
            int(round(p_value)) if p_value > 0 else "",
            int(round(a_value)) if a_value > 0 else "",
        ]
        for col_idx, value in enumerate(values, start=1):
            cell = sheet.cell(row_idx, col_idx, value)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = black_border
        row_idx += 1

    tt_row = row_idx
    tt_fill = PatternFill(fill_type="solid", fgColor="FFC0C0C0")
    tt_values = ["TT", int(round(total_m)), int(round(total_p)), int(round(total_a))]
    for col_idx, value in enumerate(tt_values, start=1):
        cell = sheet.cell(tt_row, col_idx, value)
        cell.font = Font(bold=True, size=11)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = black_border
        cell.fill = tt_fill

    # Colonne de separation
    sheet.column_dimensions[get_column_letter(5)].width = 3

    # Bloc droite: details OF retard en conservant la mise en page de Global
    details_start_col = 6
    for offset, col_name in enumerate(columns):
        target_col = details_start_col + offset
        sheet.cell(1, target_col, col_name)

    for data_row_idx, row in enumerate(retard_rows, start=2):
        for offset, col_name in enumerate(columns):
            target_col = details_start_col + offset
            sheet.cell(
                data_row_idx,
                target_col,
                _format_cell_value(col_name, row.get(col_name, ""), column_format),
            )

    details_end_row = max(1, len(retard_rows) + 1)
    if columns:
        details_end_col = details_start_col + len(columns) - 1
        table_ref = f"{get_column_letter(details_start_col)}1:{get_column_letter(details_end_col)}{details_end_row}"
        table = Table(displayName="TableRetardOF", ref=table_ref)
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        sheet.add_table(table)

    for offset, col_name in enumerate(columns):
        target_col = details_start_col + offset
        fmt = column_format.get(col_name, {})

        width = fmt.get("width")
        if width is not None:
            sheet.column_dimensions[get_column_letter(target_col)].width = float(width)

        align_value = str(fmt.get("align", "")).lower()
        horizontal = None
        if align_value in {"center", "centre"}:
            horizontal = "center"
        elif align_value in {"left", "gauche"}:
            horizontal = "left"

        bold = bool(fmt.get("bold", False))
        font_size = fmt.get("font_size")
        font_size_value = float(font_size) if font_size is not None else None

        for row_idx in range(1, details_end_row + 1):
            cell = sheet.cell(row_idx, target_col)
            if horizontal:
                cell.alignment = Alignment(horizontal=horizontal, vertical="center")
            else:
                cell.alignment = Alignment(vertical="center")
            if bold or font_size_value is not None:
                cell.font = Font(bold=bold, size=font_size_value)
            cell.border = black_border


def _week_sort_key(row: dict[str, Any]) -> tuple[int, int, int, date]:
    week_raw = row.get("SEMAINE", "")
    dt_liv = _parse_french_date(row.get("DT_LIV_CONFIRMEE"))
    date_missing = 0 if dt_liv is not None else 1
    safe_date = dt_liv or date.max

    try:
        week = int(week_raw)
        return (0, week, date_missing, safe_date)
    except (TypeError, ValueError):
        return (1, 999, date_missing, safe_date)


def _apply_sheet_format(sheet: Any, columns: list[str], column_format: dict[str, dict[str, Any]]) -> None:
    for idx, col in enumerate(columns, start=1):
        fmt = column_format.get(col, {})
        width = fmt.get("width")
        if width:
            sheet.column_dimensions[get_column_letter(idx)].width = float(width)

        align_value = str(fmt.get("align", "")).lower()
        horizontal = None
        if align_value in {"center", "centre"}:
            horizontal = "center"
        elif align_value in {"left", "gauche"}:
            horizontal = "left"

        bold = bool(fmt.get("bold", False))
        font_size = fmt.get("font_size")
        font_size_value = float(font_size) if font_size is not None else None

        for row_idx in range(1, sheet.max_row + 1):
            cell = sheet.cell(row_idx, idx)
            if horizontal:
                cell.alignment = Alignment(horizontal=horizontal)
            if bold or font_size_value is not None:
                cell.font = Font(bold=bold, size=font_size_value)


def _write_output(
    input_path: Path,
    output_path: Path,
    output_sheet: str,
    columns: list[str],
    rows: list[dict[str, Any]],
    column_format: dict[str, dict[str, Any]],
    create_excel_table: bool,
    create_weekly_summary: bool,
    summary_sheet_name: str,
) -> None:
    suffix = input_path.suffix.lower()

    if suffix in {".xlsx", ".xlsm"}:
        workbook = load_workbook(input_path, keep_vba=(suffix == ".xlsm"))
    else:
        workbook = Workbook()
        workbook.remove(workbook.active)

    if output_sheet in workbook.sheetnames:
        del workbook[output_sheet]

    sheet = workbook.create_sheet(output_sheet)
    sheet.append(columns)
    for row in rows:
        sheet.append([_format_cell_value(col, row.get(col, ""), column_format) for col in columns])

    if create_excel_table and sheet.max_row >= 1 and sheet.max_column >= 1:
        table_name = f"Table_{output_sheet}".replace(" ", "_").replace("-", "_")
        table_ref = f"A1:{get_column_letter(sheet.max_column)}{sheet.max_row}"
        table = Table(displayName=table_name[:31], ref=table_ref)
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        sheet.add_table(table)

    _apply_sheet_format(sheet, columns, column_format)

    if create_weekly_summary:
        _write_weekly_summary_sheet(workbook, summary_sheet_name, rows)
    
    _write_retard_sheet(workbook, "Retard", rows, columns, column_format)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)


def build_v1_sheet(
    input_path: Path,
    output_path: Path | None,
    profile_path: Path,
    output_sheet: str | None = None,
) -> TransformResult:
    """Genere une feuille de sortie depuis un profil JSON."""
    profile = _parse_profile(profile_path)
    headers, rows = _read_rows(input_path)
    _validate_structure(headers, profile.expected_columns)

    transformed_rows = [
        _transform_row(row, profile.computed_columns)
        for row in rows
        if _is_row_kept(row, profile.filters)
    ]

    if profile.sort_by_week and any("SEMAINE" in row for row in transformed_rows):
        transformed_rows.sort(key=_week_sort_key)

    columns = _ordered_columns(
        headers,
        transformed_rows,
        preferred_order=profile.column_order,
        drop_columns=profile.drop_columns,
    )

    sheet_name = output_sheet or profile.output_sheet
    final_output = output_path
    if final_output is None:
        if input_path.suffix.lower() in {".xlsx", ".xlsm"}:
            final_output = input_path
        else:
            final_output = input_path.with_name(f"{input_path.stem}_resultat.xlsx")

    _write_output(
        input_path,
        final_output,
        sheet_name,
        columns,
        transformed_rows,
        profile.column_format,
        profile.create_excel_table,
        profile.create_weekly_summary,
        profile.summary_sheet_name,
    )

    return TransformResult(
        rows_total=len(rows),
        rows_kept=len(transformed_rows),
        output_path=final_output,
        output_sheet=sheet_name,
    )