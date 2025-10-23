import os
import glob
import re
import datetime
import pandas as pd
from openpyxl import load_workbook, Workbook
from pathlib import Path

# Путь к файлу LE.txt
LE_FILE = "LE.txt"
OUT_DIR = "out"
SKIPPED_FILE = "skipped.xlsx"
ERRORS_FILE = "errors.xlsx"
LOG_FILE = "log.md"

# Создаём папку out, если её нет
os.makedirs(OUT_DIR, exist_ok=True)

# Очищаем каталог out/
for file in os.listdir(OUT_DIR):
    os.remove(os.path.join(OUT_DIR, file))

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
        # Логирование в log.md
        with open(LOG_FILE, "a", encoding="utf-8") as log:
            log.write(f"\n## Загруженные LE ({datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')})\n")
            log.write(f"Загружено {len(le_set)} LE из {le_file}:\n")
            for le in sorted(le_set):
                log.write(f"- {le}\n")
    except Exception as e:
        print(f"Ошибка при чтении {le_file}: {e}")
        with open(LOG_FILE, "a", encoding="utf-8") as log:
            log.write(f"\n## Ошибка загрузки LE ({datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')})\n")
            log.write(f"Ошибка при чтении {le_file}: {e}\n")
    return le_set

def is_empty_row(row) -> bool:
    """Проверяет, является ли строка полностью пустой (все ячейки None, пустые строки, пробелы или не значимые символы)."""
    for cell in row:
        if cell.value is not None:
            value_str = str(cell.value).strip()

            # Проверяем, что строка не состоит только из пробелов и не значимых символов
            if value_str and not all(c.isspace() or c in ' \t\n\r\f\v' for c in value_str):
                return False
    return True

def parse_and_convert_amount(amount_str: str) -> tuple[float | None, str]:
    """
    Обрабатывает строку суммы:
    1. Убирает разделители разрядов (например, "427,680,000.00" → "427680000.00")
    2. Преобразует в число с российским форматом (запятая как разделитель целой и дробной части)
    3. Возвращает (число, описание ошибки) или (None, описание ошибки) при ошибке
    """
    if not amount_str:
        return None, "Пустое значение суммы"
    
    # Убираем разделители разрядов: заменяем запятые на пустую строку
    cleaned = re.sub(r',', '', str(amount_str))
    
    # Проверяем, что осталось только цифры и точка (для дробной части)
    if not re.match(r'^\d+(\.\d+)?$', cleaned):
        return None, f"Некорректный формат суммы: {amount_str}"
    
    try:
        # Преобразуем в float
        amount = float(cleaned)
        
        return amount, ""
        
    except Exception as e:
        return None, f"Ошибка преобразования в число: {e}"

def write_filtered_rows(file_path: str, le_set: set, skipped_wb, errors_ws):
    """
    Записывает строки, где в строке есть "LE" (без учета регистра) и следующая ячейка (Аналитика) совпадает с LE.txt.
    Использует pandas для чтения и обработки данных, openpyxl для сохранения стилей.
    Записывает пропущенные строки в skipped.xlsx (лист по файлу).
    Записывает ошибки в errors.xlsx.
    Возвращает количество отфильтрованных, пропущенных и ошибок.
    """
    try:
        # Читаем файл с помощью pandas
        df = pd.read_excel(file_path, header=None)
    except Exception as e:
        print(f"Ошибка при чтении {file_path}: {e}")
        errors_ws.append([Path(file_path).name, "", f"Ошибка чтения файла: {e}"])
        return 0, 0, 1

    if df.empty:
        print(f"Нет данных в {file_path}")
        errors_ws.append([Path(file_path).name, "", "Нет данных в файле"])
        return 0, 0, 1

    # Определяем заголовок (первая непустая строка)
    header_row = None
    for idx in df.index:
        if df.loc[idx].notna().any():
            header_row = idx
            break

    if header_row is None:
        print(f"Нет данных в {file_path}")
        errors_ws.append([Path(file_path).name, "", "Нет данных в файле"])
        return 0, 0, 1

    # Устанавливаем заголовки
    df.columns = df.loc[header_row]
    df = df.loc[header_row + 1:].reset_index(drop=True)

    # Создаём новый файл для вывода
    out_file = os.path.join(OUT_DIR, Path(file_path).name)
    wb_out = load_workbook(file_path)
    ws_out = wb_out.active

    if not ws_out:
        print(f"False Невозможно получить активный лист из {file_path}")
        errors_ws.append([Path(file_path).name, "", "Невозможно получить активный лист"])
        return 0, 0, 1

    # Очищаем лист, оставляя заголовки
    for row in ws_out.iter_rows(min_row=2):
        for cell in row:
            cell.value = None

    # Устанавливаем заголовки в первой строке и копируем стили
    for col_idx in range(len(df.columns)):
        header_cell = ws_out.cell(row=1, column=col_idx + 1)
        orig_header_cell = ws_out.cell(row=header_row + 1, column=col_idx + 1)
        try:
            header_cell.value = df.columns[col_idx]
        except:
            # Если merged cell или другая ошибка, пропускаем
            pass
        if orig_header_cell.has_style:
            try:
                header_cell.style = orig_header_cell.style
            except Exception as e:
                print(f"⚠️ Ошибка при копировании стиля заголовка: {e}")

    file_name = Path(file_path).name
    filtered_count = 0
    skipped_count = 0
    error_count = 0
    skipped_rows = []  # Список индексов для skipped

    # Логирование начала обработки файла
    with open(LOG_FILE, "a", encoding="utf-8") as log:
        log.write(f"\n### Обработка строк файла {Path(file_path).name} ({datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')})\n")

    # Обрабатываем строки данных
    for row_idx, row in df.iterrows():
        # Пропускаем пустые строки
        if row.isna().all():
            continue

        match_found = False
        reason = ""
        is_error = False
        error_desc = ""

        # Ищем "LE" в строке
        for col_idx, (col_name, value) in enumerate(row.items()):
            cell_value = str(value).strip().upper() if pd.notna(value) else ""
            if cell_value == "LE":
                # Следующая ячейка как Аналитика
                if col_idx + 1 < len(row):
                    analytics_value = str(row.iloc[col_idx + 1]).strip().replace("-", "").replace(" ", "").upper() if pd.notna(row.iloc[col_idx + 1]) else ""
                    if analytics_value and analytics_value in le_set:
                        match_found = True
                        reason = f"Фильтрация: да (LE в колонке {col_idx}, Аналитика='{analytics_value}')"
                        break
                    elif analytics_value:
                        reason = f"Фильтрация: нет (LE в колонке {col_idx}, Аналитика='{analytics_value}' не в LE.txt)"
                    else:
                        is_error = True
                        error_desc = f"Пустое значение Аналитики после LE в колонке {col_idx}"
                else:
                    is_error = True
                    error_desc = f"LE в колонке {col_idx}, но нет следующей ячейки для Аналитики"
                break

        if not match_found and not reason and not is_error:
            is_error = True
            error_desc = "Не найдено 'LE' в строке"

        # Логирование каждой строки
        with open(LOG_FILE, "a", encoding="utf-8") as log:
            if is_error:
                log.write(f"- Строка {row_idx + header_row + 2}: Ошибка - {error_desc}\n")
            else:
                log.write(f"- Строка {row_idx + header_row + 2}: {reason}\n")

        # Обработка строк нужных данных
        if match_found:
            # Преобразование суммы проводки (предполагаем колонку 7, индекс 7)
            amount_str = str(row.iloc[7]).strip() if pd.notna(row.iloc[7]) else ""
            parsed_amount, error_msg = parse_and_convert_amount(amount_str)

            if error_msg:
                errors_ws.append([file_name, str(row_idx + header_row + 2), f"Ошибка преобразования суммы: {error_msg}"])
                is_error = True
                error_desc = f"Ошибка преобразования суммы: {error_msg}"
            else:
                # Успешное преобразование
                row.iloc[7] = parsed_amount

            # Копируем строку с сохранением форматов
            for col_idx in range(len(row)):
                out_cell = ws_out.cell(row=filtered_count + 2, column=col_idx + 1)
                out_cell.value = row.iloc[col_idx]

                # Копируем стиль из оригинального файла
                orig_cell = ws_out.cell(row=row_idx + header_row + 2, column=col_idx + 1)
                if orig_cell.has_style:
                    try:
                        out_cell.style = orig_cell.style
                    except Exception as e:
                        print(f"⚠️ Ошибка при копировании стиля: {e}")

                if col_idx == 7:  # Сумма
                    out_cell.number_format = '#,##0.00'  # стандартный Excel формат числа с 2 знаками

                if col_idx == 3:  # Дата
                    out_cell.number_format = 'dd.mm.yyyy'


            filtered_count += 1

        elif is_error:
            error_count += 1
            errors_ws.append([file_name, str(row_idx + header_row + 2), error_desc])
            skipped_rows.append(row_idx)
            skipped_count += 1
        else:
            skipped_count += 1
            skipped_rows.append(row_idx)

    # Создаём лист в skipped_wb для пропущенных строк
    if skipped_rows:
        skipped_ws = skipped_wb.create_sheet(title=file_name[:31])
        # Копируем заголовки
        for col in range(len(df.columns)):
            skipped_ws.cell(row=1, column=col + 1).value = df.columns[col]
        # Записываем пропущенные строки
        for i, row_idx in enumerate(skipped_rows):
            row = df.iloc[row_idx]
            for col_idx in range(len(row)):
                skipped_ws.cell(row=i + 2, column=col_idx + 1).value = row.iloc[col_idx]
                # Копируем стиль
                orig_cell = ws_out.cell(row=row_idx + header_row + 2, column=col_idx + 1)
                if orig_cell.has_style:
                    try:
                        skipped_ws.cell(row=i + 2, column=col_idx + 1).style = orig_cell.style
                    except Exception as e:
                        print(f"⚠️ Ошибка при копировании стиля в skipped: {e}")

    # Удаляем пустые строки в конце
    max_row = ws_out.max_row
    while max_row > 1 and all(cell.value is None or str(cell.value).strip() == "" for cell in ws_out[max_row]):
        ws_out.delete_rows(max_row)
        max_row -= 1

    # Сохраняем результат
    if filtered_count > 0:
        try:
            wb_out.save(out_file)
            print(f"Записано {filtered_count} проводок в {out_file}")
        except Exception as e:
            print(f"False Ошибка при сохранении {out_file}: {e}")
            errors_ws.append([file_name, "", f"Ошибка сохранения файла: {e}"])
            return 0, 0, error_count + 1
    else:
        print(f"Нет проводок для записи в {out_file} - файл не создан")

    # Логирование итоговых статистик
    with open(LOG_FILE, "a", encoding="utf-8") as log:
        log.write(f"\n### Итоговые статистики для файла {Path(file_path).name} ({datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')})\n")
        log.write(f"- Отфильтровано: {filtered_count} строк\n")
        log.write(f"- Пропущено: {skipped_count} строк\n")
        log.write(f"- Ошибки: {error_count} строк\n")

    return filtered_count, skipped_count, error_count

def main():
    # Логирование начала обработки
    with open(LOG_FILE, "a", encoding="utf-8") as log:
        log.write(f"\n# Запуск скрипта ({datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')})\n")
        log.write("Команда: python filter_diasoft_acc_by_LE.py\n")

    le_set = load_le_set(LE_FILE)
    if not le_set:
        print("False Нет данных в LE.txt — завершаем")
        with open(LOG_FILE, "a", encoding="utf-8") as log:
            log.write("\n## Завершение с ошибкой\n")
            log.write("Нет данных в LE.txt\n")
        return

    # Поиск файлов
    in_files = glob.glob("in/*.xlsx")
    if not in_files:
        print("False Нет файлов в папке in/")
        with open(LOG_FILE, "a", encoding="utf-8") as log:
            log.write("\n## Завершение с ошибкой\n")
            log.write("Нет файлов в папке in/\n")
        return

    print(f"Найдено {len(in_files)} файлов для обработки")
    with open(LOG_FILE, "a", encoding="utf-8") as log:
        log.write(f"\n## Найденные файлы для обработки ({datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')})\n")
        log.write(f"Найдено {len(in_files)} файлов:\n")
        for file_path in in_files:
            log.write(f"- {Path(file_path).name}\n")

    total_filtered = 0
    total_skipped = 0
    total_errors = 0

    # Очищаем skipped.xlsx и errors.xlsx перед запуском
    if os.path.exists(os.path.join(OUT_DIR, SKIPPED_FILE)):
        os.remove(os.path.join(OUT_DIR, SKIPPED_FILE))
    if os.path.exists(os.path.join(OUT_DIR, ERRORS_FILE)):
        os.remove(os.path.join(OUT_DIR, ERRORS_FILE))

    # Подготовка файла пропущенных строк
    skipped_wb = Workbook()
    # Удаляем стандартный лист
    if skipped_wb.active:
        skipped_wb.remove(skipped_wb.active)

    # Подготовка файла ошибок
    errors_wb = Workbook()
    errors_ws = errors_wb.active

    if not errors_ws:
        print(f"False Невозможно создать активный лист для errors.xlsx")
        return

    errors_ws.append(["Файл", "Строка", "Описание ошибки"])

    for file_path in in_files:
        print(f"\n{'='*50}")
        print(f"Обработка файла: {file_path}")
        print(f"{'='*50}")

        try:
            wb = load_workbook(file_path)
            ws = wb.active

            # Записываем фильтрованные строки (новая логика без пар колонок)
            filtered, skipped, errors = write_filtered_rows(file_path, le_set, skipped_wb, errors_ws)
            total_filtered += filtered
            total_skipped += skipped
            total_errors += errors

        except Exception as e:
            print(f"Ошибка при обработке {file_path}: {e}")
            errors_ws.append([Path(file_path).name, "", f"Ошибка загрузки файла: {e}"])
            total_errors += 1
            with open(LOG_FILE, "a", encoding="utf-8") as log:
                log.write(f"\n### Ошибка обработки файла {Path(file_path).name} ({datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')})\n")
                log.write(f"Ошибка: {e}\n")

    # Сохраняем skipped.xlsx после обработки всех файлов только если есть листы
    if len(skipped_wb.sheetnames) > 0:
        try:
            skipped_wb.save(os.path.join(OUT_DIR, SKIPPED_FILE))
            print(f"Всего пропущено {total_skipped} строк (пропущенные строки записаны в skipped.xlsx)")
        except Exception as e:
            print(f"False Ошибка при сохранении {SKIPPED_FILE}: {e}")

    # Сохраняем errors.xlsx после обработки всех файлов
    if total_errors > 0:
        try:
            errors_wb.save(os.path.join(OUT_DIR, ERRORS_FILE))
            print(f"Всего ошибок {total_errors} (ошибки записаны в errors.xlsx)")
        except Exception as e:
            print(f"False Ошибка при сохранении {ERRORS_FILE}: {e}")

    print("\n" + "="*50)
    print("Обработка завершена")
    print("="*50)

    # Логирование завершения
    with open(LOG_FILE, "a", encoding="utf-8") as log:
        log.write(f"\n## Обработка завершена ({datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')})\n")
        log.write("Все файлы обработаны.\n")

if __name__ == "__main__":
    main()
