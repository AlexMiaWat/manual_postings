import os
import glob
import re
import datetime
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

def write_filtered_rows(sheet, file_path: str, le_set: set, skipped_wb, errors_ws):
    """
    Записывает строки, где в строке есть "LE" (без учета регистра) и следующая ячейка (Аналитика) совпадает с LE.txt.
    Сохраняет форматы и стили исходного файла.
    Записывает пропущенные строки в skipped.xlsx (лист по файлу).
    Записывает ошибки в errors.xlsx.
    Возвращает количество отфильтрованных, пропущенных и ошибок.
    """
    first_data_row = None
    for row in sheet.iter_rows(min_row=1, max_row=20):
        if any(cell.value is not None and str(cell.value).strip() != "" for cell in row):
            first_data_row = row[0].row
            break

    if not first_data_row:
        print(f"Нет данных в {file_path}")
        errors_ws.append([Path(file_path).name, "", "Нет данных в файле"])
        return 0, 0, 1

    # Создаём новый файл для вывода только если есть данные для записи
    out_file = os.path.join(OUT_DIR, Path(file_path).name)
    wb_out = load_workbook(file_path)
    ws_out = wb_out.active

    if not ws_out:
        print(f"False Невозможно получить активный лист из {file_path}")
        errors_ws.append([Path(file_path).name, "", "Невозможно получить активный лист"])
        return 0, 0, 1

    # Удаляем старые данные (оставляем только заголовки)
    for row in ws_out.iter_rows(min_row=1):
        for cell in row:
            cell.value = None

    # Копируем все заголовки
    header_row = first_data_row
    for col in sheet.iter_cols(min_row=header_row, max_row=header_row):
        ws_out.cell(row=1, column=col[0].column).value = col[0].value
        ws_out.cell(row=1, column=col[0].column).style = col[0].style
        ws_out.cell(row=1, column=col[0].column).number_format = col[0].number_format

    file_name = Path(file_path).name
    filtered_count = 0
    skipped_count = 0
    error_count = 0
    skipped_rows = []  # Список строк для skipped

    # Логирование начала обработки файла
    with open(LOG_FILE, "a", encoding="utf-8") as log:
        log.write(f"\n### Обработка строк файла {Path(file_path).name} ({datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')})\n")

    # Обрабатываем строки данных
    for row_idx, row in enumerate(sheet.iter_rows(min_row=header_row + 1), start=header_row + 1):
        # Пропускаем пустые строки
        if is_empty_row(row):
            continue

        match_found = False
        reason = ""
        is_error = False
        error_desc = ""
        # Проходим по ячейкам в строке
        for col_idx, cell in enumerate(row, start=1):
            cell_value = str(cell.value).strip().upper() if cell.value is not None else ""
            if cell_value == "LE":
                # Следующая ячейка как Аналитика
                if col_idx < len(row):
                    analytics_cell = row[col_idx]  # col_idx уже +1 от enumerate
                    analytics_value = str(analytics_cell.value).strip().replace("-", "").replace(" ", "").upper() if analytics_cell.value is not None else ""
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
                break  # Нашли LE, не ищем дальше в строке

        if not match_found and not reason and not is_error:
            is_error = True
            error_desc = "Не найдено 'LE' в строке"

        # Логирование каждой строки
        with open(LOG_FILE, "a", encoding="utf-8") as log:
            if is_error:
                log.write(f"- Строка {row_idx}: Ошибка - {error_desc}\n")
            else:
                log.write(f"- Строка {row_idx}: {reason}\n")

        # Обработка строк нужных данных
        if match_found:
            
            # Преобразование суммы проводки
            # Предполагаем, что сумма в колонке 8 (индекс 7)
            amount_str = str(row[7].value).strip() if row[7].value is not None else ""
            parsed_amount, error_msg = parse_and_convert_amount(amount_str)
            
            if error_msg:
                # Ошибка преобразования — добавляем строку в errors.xlsx
                errors_ws.append([file_name, str(row_idx), f"Ошибка преобразования суммы: {error_msg}"])
                is_error = True
                error_desc = f"Ошибка преобразования суммы: {error_msg}"
            else:
                # Успешное преобразование — сохраняем число в ячейке
                row[7].value = parsed_amount
                row[7].number_format = '# ##0.00'

            # Копируем строку с сохранением форматов
            for col_idx, cell in enumerate(row, start=1):
                ws_out.cell(row=filtered_count + 2, column=col_idx).value = cell.value
                if col_idx == 8:
                    ws_out.cell(row=filtered_count + 2, column=col_idx).number_format = '# ##0.00'
                
                # Копируем стиль
                if cell.has_style:
                    try:
                        # Используем метод copy_style для безопасного копирования стиля
                       ws_out.cell(row=filtered_count + 2, column=col_idx).style = cell.style
                    except Exception as e:
                        print(f"⚠️ Ошибка при копировании стиля: {e}")


            filtered_count += 1

        elif is_error:
            error_count += 1
            errors_ws.append([file_name, str(row_idx), error_desc])
            # Строки с ошибками также добавляем в skipped_rows
            skipped_rows.append(row)
            skipped_count += 1
        else:
            skipped_count += 1
            # Добавляем строку в skipped_rows
            skipped_rows.append(row)

    # Создаём лист в skipped_wb только если есть данные для записи
    if skipped_rows:
        skipped_ws = skipped_wb.create_sheet(title=file_name[:31])  # Ограничение на длину имени листа
        # Копируем заголовки в skipped_ws
        for col in sheet.iter_cols(min_row=header_row, max_row=header_row):
            skipped_ws.cell(row=1, column=col[0].column).value = col[0].value
        # Записываем строки в skipped_ws
        for row_idx, row in enumerate(skipped_rows, start=1):
            for col_idx, cell in enumerate(row, start=1):
                skipped_ws.cell(row=row_idx + 1, column=col_idx).value = cell.value
                if cell.has_style:
                    try:
                        skipped_ws.cell(row=row_idx + 1, column=col_idx).style = cell.style
                    except Exception as e:
                        print(f"⚠️ Ошибка при копировании стиля в skipped: {e}")

    # Сохраняем результат только если есть отфильтрованные строки
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

    # Сохранение skipped.xlsx и errors.xlsx будет в main после обработки всех файлов

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
            filtered, skipped, errors = write_filtered_rows(ws, file_path, le_set, skipped_wb, errors_ws)
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
