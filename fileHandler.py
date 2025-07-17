from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# --- Константы ---
INPUT_FILENAME = "Расширенная+оборотно-сальдовая+ведомость+(на+проверку).xlsx"
OUTPUT_FILENAME = "outputTest.xlsx"
ALLOWED_DEVIATION_PERCENTAGE = 0.035
HEADER_ROW_INDEX = 6
START_ROW_INDEX = 8
VALUE_NORMAL = "Норма"
TURNOVER_EXCESS_PERCENT = f"Превышение {round(ALLOWED_DEVIATION_PERCENTAGE * 100, 1)}% от оборота"
SURPLUS_WITH_ZERO_TURNOVER = "Излишек при нулевом обороте"
SHORTAGE_WITH_ZERO_TURNOVER = "Недостача при нулевом обороте"


COLOR_BLACK = "000000"
COLOR_BLUE = "FFD9E1F2"
COLOR_RED = "FFF4CCCC"

COLUMN_SETTINGS = {
    "W": {"index": 23, "name": "Отклонение излишков", "fill": COLOR_BLUE},
    "X": {"index": 24, "name": "Отклонение недостач", "fill": COLOR_RED},
    "Y": {"index": 25, "name": "Норма отклонения"},
    "Z": {"index": 26, "name": "Кол-во превышения нормы"},
    "AA": {"index": 27, "name": "Сумма превышения нормы излишков"},
    "AB": {"index": 28, "name": "Сумма превышения нормы недостач"},
}

# --- Утилиты ---

def write_cell(sheet, row: int, col: int, value, fill=None):
    cell = sheet.cell(row=row, column=col, value=value)
    if fill:
        cell.fill = fill
    return cell

def create_styles():
    border = Border(
        left=Side(border_style="thin", color=COLOR_BLACK),
        right=Side(border_style="thin", color=COLOR_BLACK),
        top=Side(border_style="thin", color=COLOR_BLACK),
        bottom=Side(border_style="thin", color=COLOR_BLACK)
    )
    return border

def apply_column_widths(sheet):
    for col_data in COLUMN_SETTINGS.values():
        letter = get_column_letter(col_data["index"])
        sheet.column_dimensions[letter].width = len(col_data["name"]) + 2

def write_headers(sheet):
    for key, col_data in COLUMN_SETTINGS.items():
        fill = PatternFill(start_color=col_data.get("fill", ""), end_color=col_data.get("fill", ""), fill_type="solid") if "fill" in col_data else None
        write_cell(
            sheet,
            HEADER_ROW_INDEX,
            col_data["index"],
            col_data["name"],
            fill=fill
        )

def process_row(row_idx, sheet, border):
    # Получаем значения из колонок H, J, L, Q, S
    h = sheet.cell(row=row_idx, column=8).value or 0
    j = sheet.cell(row=row_idx, column=10).value or 0
    l = sheet.cell(row=row_idx, column=12).value or 0
    q = sheet.cell(row=row_idx, column=17).value or 0
    s = sheet.cell(row=row_idx, column=19).value or 0

    # Значения для формулы AA
    d = sheet.cell(row=row_idx, column=4).value or 0
    e = sheet.cell(row=row_idx, column=5).value or 0
    f = sheet.cell(row=row_idx, column=6).value or 0
    g = sheet.cell(row=row_idx, column=7).value or 0
    i = sheet.cell(row=row_idx, column=9).value or 0
    k = sheet.cell(row=row_idx, column=11).value or 0
    n = sheet.cell(row=row_idx, column=14).value or 0
    u = sheet.cell(row=row_idx, column=21).value or 0
    v = sheet.cell(row=row_idx, column=22).value or 0
    
    h, j, l, q, s = map(lambda x: x if isinstance(x, (int, float)) else 0, (h, j, l, q, s))
    turnover = abs(h) + abs(j) + abs(l)
    threshold = ALLOWED_DEVIATION_PERCENTAGE * turnover

    # W
    if turnover == 0 and q != 0:
        result_W = SURPLUS_WITH_ZERO_TURNOVER
    elif abs(q) > threshold:
        result_W = TURNOVER_EXCESS_PERCENT
    else:
        result_W = VALUE_NORMAL

    # X
    if turnover == 0 and s != 0:
        result_X = SHORTAGE_WITH_ZERO_TURNOVER
    elif abs(s) > threshold:
        result_X = TURNOVER_EXCESS_PERCENT
    else:
        result_X = VALUE_NORMAL

    # Y
    result_Y = threshold

    # Z
    if abs(q) > result_Y or abs(s) > result_Y:
        deviation_q = abs(q) - result_Y if abs(q) > result_Y else float("inf")
        deviation_s = abs(s) - result_Y if abs(s) > result_Y else float("inf")
        result_Z = min(deviation_q, deviation_s)
    else:
        result_Z = VALUE_NORMAL

    # AA – Сумма превышения нормы излишков
    if isinstance(result_Y, (int, float)) and abs(q) > result_Y:
        base_q = abs(q) - result_Y
        multiplier = None

        if u and isinstance(v, (int, float)):
            multiplier = v / u
        elif d and isinstance(e, (int, float)):
            multiplier = e / d
        elif f and isinstance(g, (int, float)):
            multiplier = g / f
        elif h and isinstance(i, (int, float)):
            multiplier = i / h
        elif j and isinstance(k, (int, float)):
            multiplier = k / j
        elif l and isinstance(n, (int, float)):
            multiplier = n / l

        result_AA = base_q * multiplier if multiplier is not None else ""
    else:
        result_AA = ""

    # AB – Сумма превышения нормы недостач
    if isinstance(result_Y, (int, float)) and abs(s) > result_Y:
        base_s = abs(s) - result_Y
        multiplier_s = None

        if u and isinstance(v, (int, float)):
            multiplier_s = v / u
        elif d and isinstance(e, (int, float)):
            multiplier_s = e / d
        elif f and isinstance(g, (int, float)):
            multiplier_s = g / f
        elif h and isinstance(i, (int, float)):
            multiplier_s = i / h
        elif j and isinstance(k, (int, float)):
            multiplier_s = k / j
        elif l and isinstance(n, (int, float)):
            multiplier_s = n / l

        result_AB = base_s * multiplier_s if multiplier_s is not None else ""
    else:
        result_AB = ""

    # Запись в ячейки
    write_cell(sheet, row_idx, COLUMN_SETTINGS["W"]["index"], result_W)
    write_cell(sheet, row_idx, COLUMN_SETTINGS["X"]["index"], result_X)
    write_cell(sheet, row_idx, COLUMN_SETTINGS["Y"]["index"], result_Y)
    write_cell(sheet, row_idx, COLUMN_SETTINGS["Z"]["index"], result_Z)
    write_cell(sheet, row_idx, COLUMN_SETTINGS["AA"]["index"], result_AA)
    write_cell(sheet, row_idx, COLUMN_SETTINGS["AB"]["index"], result_AB)

    # Возврат значений для дальнейшей обработки
    return result_W, result_X


def process_sheet(sheet):
    border = create_styles()
    write_headers(sheet)
    apply_column_widths(sheet)
    
    total_AA = 0.0
    total_AB = 0.0

    # Внутри process_sheet, перед циклом:
    blue_fill = PatternFill(start_color=COLOR_BLUE, end_color=COLOR_BLUE, fill_type="solid")
    red_fill = PatternFill(start_color=COLOR_RED, end_color=COLOR_RED, fill_type="solid")

    for row in range(START_ROW_INDEX, sheet.max_row):
        result_W, result_X = process_row(row, sheet, border)

        value_AA = sheet.cell(row=row, column=COLUMN_SETTINGS["AA"]["index"]).value
        value_AB = sheet.cell(row=row, column=COLUMN_SETTINGS["AB"]["index"]).value

        if isinstance(value_AA, (int, float)):
            total_AA += value_AA
        if isinstance(value_AB, (int, float)):
            total_AB += value_AB

        # Проверяем W
        if result_W != VALUE_NORMAL:
            for col in range(1, sheet.max_column + 1):
                sheet.cell(row=row, column=col).fill = blue_fill 

        # Проверяем X
        elif result_X != VALUE_NORMAL:
            for col in range(1, sheet.max_column + 1):
                sheet.cell(row=row, column=col).fill = red_fill 

    write_cell(sheet, sheet.max_row, COLUMN_SETTINGS["AA"]["index"], total_AA)
    write_cell(sheet, sheet.max_row, COLUMN_SETTINGS["AB"]["index"], total_AB)

    for row in range(HEADER_ROW_INDEX, sheet.max_row + 1):
        for col in range(1, sheet.max_column + 1):
            cell = sheet.cell(row=row, column=col)
            cell.border = border

def main():
    workbook = load_workbook(INPUT_FILENAME)
    sheet = workbook.worksheets[0]
    process_sheet(sheet)
    workbook.save(OUTPUT_FILENAME)

if __name__ == "__main__":
    main()
