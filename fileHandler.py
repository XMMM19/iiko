# pip install pandas xlrd==2.0.1 xlwt
import pandas as pd
import numpy as np
import xlwt
import re
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, PatternFill

# --- нормализация одного значения ---
_space_re = re.compile(r"[\u0020\u00A0\u2000-\u200B\u202F\u205F]")

def to_number(val):
    """Преобразует строки формата "2 335,99"/"1\u202F234,50"/"-7\u00A0710,11" -> float.
    Если значение не похоже на число, возвращает исходное значение без изменений.
    Поддержка скобок для отрицательных: (1\u00A0234,50) -> -1234.5
    Поддержка юникод-минуса (\u2212).
    """
    if isinstance(val, float) or isinstance(val, int):
        return float(val)
    if val is None:
        return np.nan

    s = str(val).strip()
    if s == "" or s.lower() in {"nan", "none", "null"}:
        return np.nan

    # Юникод-минус -> обычный минус
    s = s.replace("\u2212", "-")

    # Отрицательные в скобках: (123,45) -> -123,45
    is_neg = False
    if s.startswith("(") and s.endswith(")"):
        is_neg = True
        s = s[1:-1].strip()

    # Удаляем все виды пробелов/разделителей тысяч
    s_nospace = _space_re.sub("", s)

    # Заменяем запятую как десятичный разделитель на точку
    s_dot = s_nospace.replace(",", ".")

    # Разрешаем только форматы вида [+|-]digits[.digits]
    if re.fullmatch(r"[+-]?\d+(?:\.\d+)?", s_dot or ""):
        try:
            num = float(s_dot)
            return -num if is_neg else num
        except ValueError:
            return val
    else:
        return val
    

# читаем .xls
df = pd.read_excel("dev-files/in.xls", sheet_name=0, dtype=str, engine="xlrd")

# --- Удаляем строки, где в ЛЮБОЙ ячейке встречаются слова "Товар" или "кол-во" (целые слова, без учета регистра) ---
pattern = r"(?i)\b(товар|кол-во)\b"  # (?i) = ignore case, \b = границы слова
mask_has_words = df.apply(lambda s: s.astype(str).str.contains(pattern, regex=True, na=False)).any(axis=1)

# >>> Защищаем строки 7..9 от удаления <<<
# Если вы имеете в виду НОМЕРА СТРОК В EXCEL (1-based), используйте индексы 6,7,8 в pandas:
protected_rows = [7, 8, 9]  # поменяйте на [7, 8, 9], если имелись в виду индексы pandas (0-based)

# Удаляем только те, что совпали по маске И НЕ входят в защищённые индексы
mask_to_drop = mask_has_words & ~df.index.isin(protected_rows)
df = df.loc[~mask_to_drop].reset_index(drop=True)

# конвертируем все ячейки по месту (можно ограничить списком колонок)
df = df.applymap(to_number)

# --- Вписываем подписи в строку 7 (pandas) — это 9-я строка Excel ---
# Нужные колонки: 21..26 (V..AA в Excel)
needed_last_col_index = 26  # включительно
if df.shape[1] - 1 < needed_last_col_index:
    for _ in range(needed_last_col_index + 1 - df.shape[1]):
        df[f"Unnamed_{df.shape[1]}"] = pd.NA

labels = [
    (7, 21, "Отклонение изишков"),                 # V9
    (7, 22, "Отклонение недостач"),                # W9
    (7, 23, "Норма отклонения"),                   # X9
    (7, 24, "Кол-во превышения нормы"),            # Y9
    (7, 25, "Сумма превышения нормы излишков"),    # Z9
    (7, 26, "Сумма превышения нормы недостач"),    # AA9
]
for r, c, val in labels:
    df.iat[r, c] = val  # без поднятия шапок и без переименования колонок


# --- Индексы столбцов (pandas 0-based) ---
A_IDX = 0  # первая колонка — проверяем на число
J_IDX, K_IDX = 9, 10
L_IDX, M_IDX = 11, 12
N_IDX, O_IDX = 13, 14
P_IDX, R_IDX = 15, 17
T_IDX, U_IDX = 19, 20
V_IDX, W_IDX, X_IDX, Y_IDX, Z_IDX, AA_IDX = 21, 22, 23, 24, 25, 26

EPS = 1e-9

def is_number(x) -> bool:
    return isinstance(x, (int, float)) and not (isinstance(x, float) and np.isnan(x))

def nzf(x: object) -> float:
    try:
        if x is None or (isinstance(x, float) and np.isnan(x)):
            return 0.0
        return float(x)
    except Exception:
        return 0.0

for i in range(df.shape[0]):
    if i == 7:  # пропускаем строку с подписями
        continue

    first_val = df.iat[i, A_IDX] if A_IDX < df.shape[1] else None
    if not is_number(first_val):
        continue  # считаем только строки, где первая колонка — число

    # исходные значения
    j = nzf(df.iat[i, J_IDX]) if J_IDX < df.shape[1] else 0.0
    l = nzf(df.iat[i, L_IDX]) if L_IDX < df.shape[1] else 0.0
    n = nzf(df.iat[i, N_IDX]) if N_IDX < df.shape[1] else 0.0
    p = nzf(df.iat[i, P_IDX]) if P_IDX < df.shape[1] else 0.0
    r = nzf(df.iat[i, R_IDX]) if R_IDX < df.shape[1] else 0.0
    t = nzf(df.iat[i, T_IDX]) if T_IDX < df.shape[1] else 0.0
    u = df.iat[i, U_IDX] if U_IDX < df.shape[1] else None
    f = nzf(df.iat[i, 5])  if 5  < df.shape[1] else 0.0  # F
    g = df.iat[i, 6]  if 6  < df.shape[1] else None      # G
    h = nzf(df.iat[i, 7])  if 7  < df.shape[1] else 0.0  # H
    ii= df.iat[i, 8]  if 8  < df.shape[1] else None      # I
    k = df.iat[i, K_IDX] if K_IDX < df.shape[1] else None
    m = df.iat[i, M_IDX] if M_IDX < df.shape[1] else None
    o = df.iat[i, O_IDX] if O_IDX < df.shape[1] else None

    # оборот и X
    turnover = abs(j) + abs(l) + abs(n)
    x_val = 0.035 * turnover
    df.iat[i, X_IDX] = x_val

    # V по P
    if turnover <= EPS and abs(p) > EPS:
        df.iat[i, V_IDX] = "Излишек при нулевом обороте"
    elif abs(p) > x_val:
        df.iat[i, V_IDX] = "Превышение 3,5% от оборота"
    else:
        df.iat[i, V_IDX] = "Норма"

    # W по R
    if turnover <= EPS and abs(r) > EPS:
        df.iat[i, W_IDX] = "Недостача при нулевом обороте"
    elif abs(r) > x_val:
        df.iat[i, W_IDX] = "Превышение 3,5% от оборота"
    else:
        df.iat[i, W_IDX] = "Норма"

    # Y = MIN exceedance for P/R vs X, иначе "Норма"
    cond_p = abs(p) > x_val
    cond_r = abs(r) > x_val
    if cond_p or cond_r:
        d_p = abs(p) - x_val if cond_p else float("inf")
        d_r = abs(r) - x_val if cond_r else float("inf")
        df.iat[i, Y_IDX] = min(d_p, d_r)
    else:
        df.iat[i, Y_IDX] = "Норма"

    # Z = ЕСЛИ(|P|>X; (|P|-X) * tiered_ratio; "")
    z_val = ""
    if abs(p) > x_val:
        ratio = None
        # ЕСЛИ(И(T<>0; ЕЧИСЛО(U)); ABS(U)/ABS(T);
        if abs(t) > EPS and is_number(u):
            ratio = abs(float(u)) / abs(t)
        # Иначе ЕСЛИ(И(F<>0; ЕЧИСЛО(G)); G/F;
        elif abs(f) > EPS and is_number(g):
            ratio = float(g) / f
        # Иначе ЕСЛИ(И(H<>0; ЕЧИСЛО(I)); ABS(I)/ABS(H);
        elif abs(h) > EPS and is_number(ii):
            ratio = abs(float(ii)) / abs(h)
        # Иначе ЕСЛИ(И(J<>0; ЕЧИСЛО(K)); ABS(K)/ABS(J);
        elif abs(j) > EPS and is_number(k):
            ratio = abs(float(k)) / abs(j)
        # Иначе ЕСЛИ(И(L<>0; ЕЧИСЛО(M)); ABS(M)/ABS(L);
        elif abs(l) > EPS and is_number(m):
            ratio = abs(float(m)) / abs(l)
        # Иначе ЕСЛИ(И(N<>0; ЕЧИСЛО(O)); ABS(O)/ABS(N); "")
        elif abs(n) > EPS and is_number(o):
            ratio = abs(float(o)) / abs(n)

        if ratio is not None:
            z_val = (abs(p) - x_val) * ratio
        else:
            z_val = ""

    df.iat[i, Z_IDX] = z_val

    # AA = ЕСЛИ(|R|>X; (|R|-X) * tiered_ratio; "")
    aa_val = ""
    if abs(r) > x_val:
        ratio = None
        # 1) И(T<>0; ЕЧИСЛО(U)) → |U|/|T|
        if abs(t) > EPS and is_number(u):
            ratio = abs(float(u)) / abs(t)
        # 2) И(F<>0; ЕЧИСЛО(G)) → G/F
        elif abs(f) > EPS and is_number(g):
            ratio = float(g) / f
        # 3) И(H<>0; ЕЧИСЛО(I)) → |I|/|H|
        elif abs(h) > EPS and is_number(ii):
            ratio = abs(float(ii)) / abs(h)
        # 4) И(J<>0; ЕЧИСЛО(K)) → |K|/|J|
        elif abs(j) > EPS and is_number(k):
            ratio = abs(float(k)) / abs(j)
        # 5) И(L<>0; ЕЧИСЛО(M)) → |M|/|L|
        elif abs(l) > EPS and is_number(m):
            ratio = abs(float(m)) / abs(l)
        # 6) И(N<>0; ЕЧИСЛО(O)) → |O|/|N|
        elif abs(n) > EPS and is_number(o):
            ratio = abs(float(o)) / abs(n)

        if ratio is not None:
            aa_val = (abs(r) - x_val) * ratio
        else:
            aa_val = ""

    df.iat[i, AA_IDX] = aa_val


# --- Сохраняем и растягиваем колонки V..AA + перенос по словам ---
# Примечание: при header=True (по умолч.), pandas-строка 7 окажется на 9-й строке Excel.
# Вычислим номер строки в Excel, если потребуется трогать конкретную строку:
header_written = True
startrow = 0
row_in_df = 7
excel_label_row = startrow + (1 if header_written else 0) + row_in_df + 1  # = 9

with pd.ExcelWriter("dev-files/out.xlsx", engine="openpyxl") as writer:
    df.to_excel(writer, index=False, startrow=startrow)
    ws = writer.book.active

    # Ширина колонок V..AA (22..27) и перенос текста
    for col_idx in range(22, 28):  # 22=V, 27=AA
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = 35  # подберите ширину (например, 30–45)
        # Включим перенос хотя бы для строки с нашими подписями
        ws[f"{col_letter}{excel_label_row}"].alignment = Alignment(wrap_text=True, vertical="top")

    # Цвета для V9 и W9
    fill_v = PatternFill(start_color="FFD9E1F2", end_color="FFD9E1F2", fill_type="solid")  # V9
    fill_w = PatternFill(start_color="FFF4CCCC", end_color="FFF4CCCC", fill_type="solid")  # W9
    ws[f"V{excel_label_row}"].fill = fill_v
    ws[f"W{excel_label_row}"].fill = fill_w
    