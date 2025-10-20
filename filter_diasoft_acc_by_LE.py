import os
import glob
from openpyxl import load_workbook, Workbook
from pathlib import Path

# Путь к файлу LE.txt
LE_FILE = "LE.txt"
OUT_DIR = "out"
ERROR_FILE = "error.xlsx"

# Создаём папку out, если её нет
os.makedirs(OUT_DIR, exist_ok=True)


def load_le_set(le_file: str) -> set:
    """Загружает список LE из файла, убирает пробелы, тире, приводит к верхнему регистру."""
    le_set = set()
    try:
        with open(le_file, "r", encoding="utf-8") as f:
            for line in f:
                cleaned = line.strip().replace("-", "").replace(" ", "").upper()
                if cleaned:
                    le_set.add(cleaned)
        print(f"Загружено {len(le_set)} LE из {le_file}")
    except Exception as e:
        print(f"Ошибка при чтении {le_file}: {e}")
    return le_set


def find_le_columns(sheet) -> tuple[int, int]:
    """
    Находит колонку "ТИП" (с индексом: Тип1, Тип2 и т.д.) с значением "LE"
    и следующую за ней колонку "Аналитика".
    Возвращает (тип_col_idx, аналитика_col_idx), или (-1, -1) при неудаче.
    """
    first_data_row = None
    for row in sheet.iter_rows(min_row=1, max_row=20):
        if any(cell.value is not None and str(cell.value).strip() != "" for cell in row):
            first_data_row = row[0].row
            break

    if not first_data_row:
        print("Нет данных в листе")
        return -1, -1

    print(f"Первая строка с данными: {first_data_row}")

    # Ищем колонку, где значение = "LE", и название содержит "ТИП"
    type_col_idx = -1
    for col in sheet.iter_cols(min_row=first_data_row, max_row=first_data_row):
        cell_value = str(col[0].value).strip().upper()
        col_name = str(sheet.cell(row=first_data_row, column=col[0].column).column_letter).lower()

        # Добавляем отладочную информацию
        print(f"Проверка колонки {col[0].column} ({col_name}): значение='{cell_value}', содержит 'ТИП'={('ТИП' in col_name)}")

        # Проверяем, что значение — "LE" и название колонки содержит "ТИП"
        if cell_value.find("ТИП") == 0:
            type_col_idx = col[0].column
            print(f"[OK] Найдена колонка 'ТИП' с значением 'LE': {col[0].column} ({col_name})")
            break

    if type_col_idx == -1:
        print("False Не удалось найти колонку с 'ТИП' и значением 'LE'")
        return -1, -1

    # Находим следующую колонку после "ТИП"
    analytics_col_idx = type_col_idx + 1
    if analytics_col_idx > sheet.max_column:
        print(f"Нет колонки после 'ТИП' (колонка {type_col_idx})")
        return -1, -1

    # Проверяем, что следующая колонка — "Аналитика"
    analytics_cell = sheet.cell(row=first_data_row, column=analytics_col_idx)
    if not str(analytics_cell.value).strip().lower() == "аналитика":
        print(f"Следующая после 'ТИП' колонка не 'Аналитика', а: {analytics_cell.value}")
        return -1, -1

    print(f"Нашли: ТИП={type_col_idx}, Аналитика={analytics_col_idx}")
    return type_col_idx, analytics_col_idx


def write_filtered_rows(sheet, file_path: str, le_set: set, type_col_idx: int, analytics_col_idx: int):
    """
    Записывает строки, где аналитика в LE.txt.
    Сохраняет форматы и стили исходного файла.
    Записывает ошибки в error.xlsx.
    """
    first_data_row = None
    for row in sheet.iter_rows(min_row=1, max_row=20):
        if any(cell.value is not None and str(cell.value).strip() != "" for cell in row):
            first_data_row = row[0].row
            break

    if not first_data_row:
        print(f"Нет данных в {file_path}")
        return

    # Создаём новый файл для вывода
    out_file = os.path.join(OUT_DIR, Path(file_path).name)
    wb_out = load_workbook(file_path)
    ws_out = wb_out.active

    if not ws_out:
        print(f"False Невозможно получить активный лист из {file_path}")
        return

    # Удаляем старые данные (оставляем только заголовки)
    for row in ws_out.iter_rows(min_row=1):
        for cell in row:
            cell.value = None

    # Копируем заголовки
    header_row = first_data_row
    for col in sheet.iter_cols(min_row=header_row, max_row=header_row):
        if col[0].column == type_col_idx or col[0].column == analytics_col_idx:
            ws_out.cell(row=1, column=col[0].column).value = col[0].value

    # Подготовка файла ошибок
    error_wb = Workbook()
    error_ws = error_wb.active

    if not error_ws:
        print(f"False Невозможно создать активный лист для error.xlsx")
        return

    error_ws.append(["Файл", "Строка", "Описание ошибки"])

    filtered_count = 0
    error_count = 0

    # Обрабатываем строки данных
    for row_idx, row in enumerate(sheet.iter_rows(min_row=header_row + 1), start=header_row + 1):
        type_cell = row[type_col_idx - 1] if type_col_idx > 0 else None
        analytics_cell = row[analytics_col_idx - 1] if analytics_col_idx > 0 else None

        if not type_cell or not analytics_cell:
            error_count += 1
            error_ws.append([Path(file_path).name, str(row_idx), "Нет данных в колонках 'ТИП' или 'Аналитика'"])
            continue

        type_value = str(type_cell.value).strip().upper() if type_cell.value else ""
        analytics_value = str(analytics_cell.value).strip().replace("-", "").replace(" ", "").upper() if analytics_cell.value else ""

        # Проверяем, что ТИП = "LE"
        if type_value != "LE":
            error_count += 1
            error_ws.append([Path(file_path).name, str(row_idx), f"ТИП не 'LE', а: {type_value}"])
            continue

        # Сравниваем аналитику
        if analytics_value in le_set:
            # Копируем строку с сохранением форматов
            for col_idx, cell in enumerate(row, start=1):
                ws_out.cell(row=filtered_count + 2, column=col_idx).value = cell.value
                # Копируем стиль
                if cell.has_style:
                    try:
                        # Используем метод copy_style для безопасного копирования стиля
                        ws_out.cell(row=filtered_count + 2, column=col_idx).style = cell.style
                    except Exception as e:
                        print(f"⚠️ Ошибка при копировании стиля: {e}")
            filtered_count += 1
        else:
            error_count += 1
            error_ws.append([Path(file_path).name, str(row_idx), f"Аналитика '{analytics_value}' не в LE.txt"])

    # Сохраняем результат
    try:
        wb_out.save(out_file)
        print(f"Записано {filtered_count} проводок в {out_file}")
    except Exception as e:
        print(f"False Ошибка при сохранении {out_file}: {e}")
        return

    if error_count > 0:
        try:
            error_wb.save(ERROR_FILE)
            print(f"Пропущено {error_count} строк (ошибки записаны в error.xlsx)")
        except Exception as e:
            print(f"False Ошибка при сохранении {ERROR_FILE}: {e}")
            return


def main():
    le_set = load_le_set(LE_FILE)
    if not le_set:
        print("False Нет данных в LE.txt — завершаем")
        return

    # Поиск файлов
    in_files = glob.glob("in/*.xlsx")
    if not in_files:
        print("False Нет файлов в папке in/")
        return

    print(f"Найдено {len(in_files)} файлов для обработки")

    for file_path in in_files:
        print(f"\n{'='*50}")
        print(f"Обработка файла: {file_path}")
        print(f"{'='*50}")

        try:
            wb = load_workbook(file_path)
            ws = wb.active

            type_col, analytics_col = find_le_columns(ws)

            if type_col == -1 or analytics_col == -1:
                print("False Пропускаем файл — не найдены колонки 'ТИП' с 'LE' и 'Аналитика'")
                continue

            # Записываем фильтрованные строки
            write_filtered_rows(ws, file_path, le_set, type_col, analytics_col)

        except Exception as e:
            print(f"Ошибка при обработке {file_path}: {e}")

    print("\n" + "="*50)
    print("Обработка завершена")
    print("="*50)


if __name__ == "__main__":
    main()