import os
import datetime
import re


def clean_path(p):
    """Убирает лишние пробелы и кавычки вокруг пути"""
    return p.strip().strip('\'"')


def parse_date(s):
    s = s.strip().replace(" ", "")
    s = re.sub(r"[.\-\/]", "", s)
    if len(s) != 8 or not s.isdigit():
        raise ValueError("Неверный формат даты (нужно ГГГГММДД)")
    y, m, d = int(s[:4]), int(s[4:6]), int(s[6:8])
    datetime.date(y, m, d)  # проверка валидности
    return f"{y:04d}.{m:02d}.{d:02d}"


def get_date():
    while True:
        print("\nВыбор даты:")
        print("1 - Сегодня")
        print("2 - Своя (например 20251122)")
        choice = input("Введите 1 или 2: ").strip()
        
        if choice == "1":
            t = datetime.date.today()
            return t.strftime("%Y.%m.%d")
        elif choice == "2":
            raw = input("Введите дату: ")
            try:
                return parse_date(raw)
            except Exception as e:
                print("Ошибка:", e)
        else:
            print("Нужно ввести 1 или 2")


def variant_1(root, date_prefix):
    # Папка с подпапками -> файлы внутри подпапок
    root = clean_path(root)
    
    if not os.path.isdir(root):
        print(f"Ошибка: '{root}' — не папка или не найдена")
        return

    cnt = 0
    print(f"\nОбработка папки: {root}")
    
    for sub in os.listdir(root):
        sub_path = os.path.join(root, sub)
        if not os.path.isdir(sub_path):
            continue
            
        for fname in os.listdir(sub_path):
            fpath = os.path.join(sub_path, fname)
            if not os.path.isfile(fpath):
                continue
                
            name, ext = os.path.splitext(fname)
            # Проверка, чтобы не добавлять префикс повторно (опционально)
            if fname.startswith(date_prefix):
                continue
                
            new_name = f"{date_prefix}_{name}_{sub}{ext}"
            new_path = os.path.join(sub_path, new_name)
            
            try:
                os.rename(fpath, new_path)
                print(f"{fname} -> {new_name}")
                cnt += 1
            except Exception as e:
                print(f"Ошибка переименования {fname}: {e}")

    print(f"\nВсего переименовано: {cnt}")


def variant_2(target, date_prefix):
    # Файл или папка -> добавить префикс
    target = clean_path(target)

    if os.path.isfile(target):
        # Единичный файл
        folder = os.path.dirname(target) or "."
        fname = os.path.basename(target)
        new_name = f"{date_prefix}_{fname}"
        new_path = os.path.join(folder, new_name)
        
        try:
            os.rename(target, new_path)
            print(f"\n{fname} -> {new_name}")
            print("Готово (1 файл)")
        except Exception as e:
            print(f"Ошибка: {e}")

    elif os.path.isdir(target):
        # Папка — переименовать все файлы внутри
        print(f"\nОбработка файлов в папке: {target}")
        cnt = 0
        for fname in os.listdir(target):
            fpath = os.path.join(target, fname)
            if not os.path.isfile(fpath):
                continue
                
            new_name = f"{date_prefix}_{fname}"
            new_path = os.path.join(target, new_name)
            
            try:
                os.rename(fpath, new_path)
                print(f"{fname} -> {new_name}")
                cnt += 1
            except Exception as e:
                print(f"Ошибка переименования {fname}: {e}")
        print(f"\nВсего переименовано: {cnt}")
        
    else:
        print(f"Ошибка: путь '{target}' не найден")


def main():
    while True:
        print("\n=== ПЕРЕИМЕНОВАНИЕ ФАЙЛОВ ===")
        print("1 - Game (подпапки + суффикс папки)")
        print("2 - Добавить дату (файл или папка)")
        print("q - Выход")
        
        mode = input("Ваш выбор: ").strip().lower()
        
        if mode == 'q':
            break
            
        if mode not in ['1', '2']:
            print("Неверный выбор")
            continue

        try:
            date_prefix = get_date()
            print(f"Выбрана дата: {date_prefix}")
            
            if mode == "1":
                root = input("Путь к общей папке: ")
                variant_1(root, date_prefix)
            elif mode == "2":
                target = input("Путь к файлу или папке: ")
                variant_2(target, date_prefix)
                
        except Exception as e:
            print(f"Произошла ошибка: {e}")
            
        input("\nНажмите Enter, чтобы продолжить...")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nВыход...")
    except Exception as e:
        print(f"\nКритическая ошибка: {e}")
        input("Нажмите Enter для закрытия...")
