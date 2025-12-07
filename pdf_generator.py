#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт для генерации PDF-документов из CSV/JSON данных и HTML-шаблонов.
"""

import os
import json
import csv
import sys
import re
from pathlib import Path
from typing import Dict, List, Any, Optional
import pandas as pd
from xhtml2pdf import pisa
from io import BytesIO
import platform

# Настройка шрифтов для поддержки кириллицы
CYRILLIC_FONT_PATH = None
try:
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    
    # Пытаемся найти и зарегистрировать стандартные шрифты, поддерживающие кириллицу
    if platform.system() == 'Windows':
        # Стандартные пути к шрифтам Windows
        font_paths = [
            r'C:\Windows\Fonts\arial.ttf',
            r'C:\Windows\Fonts\ARIAL.TTF',
        ]
        for arial_font_path in font_paths:
            if os.path.exists(arial_font_path):
                try:
                    # Регистрируем шрифт с правильным именем
                    # Регистрируем под разными именами для максимальной совместимости
                    pdfmetrics.registerFont(TTFont('Arial', arial_font_path))
                    pdfmetrics.registerFont(TTFont('ArialUnicode', arial_font_path))
                    pdfmetrics.registerFont(TTFont('sans-serif', arial_font_path))
                    # Также регистрируем как Helvetica (стандартный шрифт reportlab)
                    pdfmetrics.registerFont(TTFont('Helvetica', arial_font_path))
                    CYRILLIC_FONT_PATH = arial_font_path
                    print("[INFO] Шрифт Arial зарегистрирован для поддержки кириллицы")
                    print(f"[INFO] Путь к шрифту: {arial_font_path}")
                    break
                except Exception as e:
                    print(f"[INFO] Не удалось зарегистрировать шрифт Arial: {e}")
    elif platform.system() == 'Darwin':  # macOS
        # Попробуем найти Arial на macOS
        arial_paths = [
            '/Library/Fonts/Arial.ttf',
            '/System/Library/Fonts/Supplemental/Arial.ttf',
        ]
        for arial_path in arial_paths:
            if os.path.exists(arial_path):
                try:
                    pdfmetrics.registerFont(TTFont('Arial', arial_path))
                    CYRILLIC_FONT_PATH = arial_path
                    break
                except:
                    pass
except ImportError:
    pass
except Exception:
    pass

# Пути к директориям
BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / 'data'
TEMPLATES_DIR = BASE_DIR / 'templates'
OUTPUT_DIR = BASE_DIR / 'output'

# Создание директорий, если их нет
DATA_DIR.mkdir(exist_ok=True)
TEMPLATES_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)


def get_data_files() -> List[Path]:
    """Получает список всех CSV и JSON файлов из директории /data."""
    data_files = []
    if DATA_DIR.exists():
        data_files.extend(DATA_DIR.glob('*.csv'))
        data_files.extend(DATA_DIR.glob('*.json'))
    return sorted(data_files)


def get_template_files() -> List[Path]:
    """Получает список всех HTML-шаблонов из директории /templates."""
    template_files = []
    if TEMPLATES_DIR.exists():
        template_files.extend(TEMPLATES_DIR.glob('*.html'))
    return sorted(template_files)


def load_csv_data(file_path: Path) -> List[Dict[str, Any]]:
    """Загружает данные из CSV файла."""
    try:
        df = pd.read_csv(file_path, encoding='utf-8')
        # Преобразуем DataFrame в список словарей, очищая пробелы в ключах
        records = df.to_dict('records')
        # Очищаем ключи от пробелов и нормализуем
        normalized_records = []
        for record in records:
            normalized = {}
            for key, value in record.items():
                # Убираем пробелы в начале и конце ключа
                clean_key = str(key).strip()
                normalized[clean_key] = value
            normalized_records.append(normalized)
        return normalized_records
    except Exception as e:
        print(f"Ошибка при чтении CSV файла через pandas: {e}")
        # Попробуем альтернативный способ через csv
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                records = list(reader)
                # Нормализуем ключи
                normalized_records = []
                for record in records:
                    normalized = {}
                    for key, value in record.items():
                        clean_key = str(key).strip()
                        normalized[clean_key] = value
                    normalized_records.append(normalized)
                return normalized_records
        except Exception as e2:
            print(f"Ошибка при чтении CSV через стандартную библиотеку: {e2}")
            return []


def load_json_data(file_path: Path) -> List[Dict[str, Any]]:
    """Загружает данные из JSON файла."""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
            # Если это список, возвращаем как есть
            if isinstance(data, list):
                return data
            # Если это словарь, оборачиваем в список
            elif isinstance(data, dict):
                return [data]
            else:
                return []
    except Exception as e:
        print(f"Ошибка при чтении JSON файла: {e}")
        return []


def load_data_file(file_path: Path) -> List[Dict[str, Any]]:
    """Загружает данные из файла (CSV или JSON)."""
    if file_path.suffix.lower() == '.csv':
        return load_csv_data(file_path)
    elif file_path.suffix.lower() == '.json':
        return load_json_data(file_path)
    else:
        return []


def get_invoice_ids(data: List[Dict[str, Any]]) -> List[str]:
    """Извлекает список уникальных invoice id из данных."""
    invoice_ids = []
    for record in data:
        # Пробуем разные варианты названий поля
        invoice_id = (
            record.get('invoice_id') or
            record.get('invoiceId') or
            record.get('invoice') or
            record.get('id') or
            record.get('ID')
        )
        if invoice_id and str(invoice_id) not in invoice_ids:
            invoice_ids.append(str(invoice_id))
    return sorted(invoice_ids)


def find_record_by_invoice_id(data: List[Dict[str, Any]], invoice_id: str) -> Optional[Dict[str, Any]]:
    """Находит запись по invoice id."""
    for record in data:
        invoice_id_field = (
            record.get('invoice_id') or
            record.get('invoiceId') or
            record.get('invoice') or
            record.get('id') or
            record.get('ID')
        )
        if str(invoice_id_field) == str(invoice_id):
            return record
    return None


def render_template(template_path: Path, data: Dict[str, Any]) -> str:
    """Подставляет данные в HTML-шаблон."""
    try:
        with open(template_path, 'r', encoding='utf-8') as f:
            template = f.read()
        
        # Подготовка данных для подстановки
        template_data = {}
        
        # Копируем все данные, преобразуя значения в строки
        for key, value in data.items():
            # Преобразуем значение в строку, обрабатывая None и NaN
            if value is None:
                template_data[key] = ''
            elif isinstance(value, float) and (value != value):  # NaN check
                template_data[key] = ''
            else:
                template_data[key] = str(value)
        
        # Обработка опциональных полей для специальных плейсхолдеров
        if 'tax' in data and data.get('tax') is not None:
            tax_value = str(data.get('tax'))
            template_data['tax_row'] = f'<div class="amount-row"><span>НДС:</span><span>{tax_value} руб.</span></div>'
        else:
            template_data['tax_row'] = ''
        
        if 'total' in data and data.get('total') is not None:
            total_value = str(data.get('total'))
            template_data['total_row'] = f'<div class="amount-row total"><span>Итого:</span><span>{total_value} руб.</span></div>'
        else:
            template_data['total_row'] = ''
        
        # Используем регулярные выражения для замены только известных плейсхолдеров
        # Это позволяет избежать проблем с фигурными скобками в CSS
        rendered = template
        
        # Сначала обрабатываем экранированные скобки {{ и }}
        # Временно заменяем их на маркеры
        rendered = re.sub(r'\{\{', '\x00OPEN\x00', rendered)
        rendered = re.sub(r'\}\}', '\x00CLOSE\x00', rendered)
        
        # Заменяем только известные плейсхолдеры
        for key, value in template_data.items():
            # Ищем плейсхолдеры вида {key} или {key:format}
            pattern = r'\{' + re.escape(key) + r'(?::[^}]*)?\}'
            rendered = re.sub(pattern, lambda m: str(value), rendered)
        
        # Возвращаем экранированные скобки обратно
        rendered = rendered.replace('\x00OPEN\x00', '{{')
        rendered = rendered.replace('\x00CLOSE\x00', '}}')
        
        # Проверяем, остались ли необработанные плейсхолдеры (только те, что не в CSS)
        # Игнорируем фигурные скобки внутри тегов <style> и CSS-селекторов
        remaining_placeholders = []
        # Ищем плейсхолдеры вне CSS-блоков
        style_pattern = r'<style[^>]*>.*?</style>'
        non_style_content = re.sub(style_pattern, '', rendered, flags=re.DOTALL | re.IGNORECASE)
        remaining = re.findall(r'\{([a-zA-Z_][a-zA-Z0-9_]*)\}', non_style_content)
        if remaining:
            remaining_placeholders = [p for p in set(remaining) if p not in template_data]
        
        if remaining_placeholders:
            print(f"\n[ВНИМАНИЕ] В шаблоне остались необработанные плейсхолдеры: {set(remaining_placeholders)}")
            print(f"[DEBUG] Доступные ключи в данных: {list(template_data.keys())}")
        
        return rendered
    except Exception as e:
        print(f"Ошибка при рендеринге шаблона: {e}")
        import traceback
        traceback.print_exc()
        return template


def generate_pdf(html_content: str, output_path: Path):
    """Генерирует PDF из HTML-контента с поддержкой кириллицы."""
    try:
        # Добавляем @font-face для явного указания шрифта, если он найден
        font_face_css = ''
        if CYRILLIC_FONT_PATH and os.path.exists(CYRILLIC_FONT_PATH):
            # xhtml2pdf требует абсолютный путь к файлу шрифта
            # Для Windows нужно использовать правильный формат пути
            if platform.system() == 'Windows':
                # Убираем ведущий слеш и используем правильный формат для Windows
                font_path_escaped = CYRILLIC_FONT_PATH.replace('\\', '/').replace('C:/', '/C/')
                font_url = f"file://{font_path_escaped}"
            else:
                font_path_escaped = CYRILLIC_FONT_PATH.replace('\\', '/')
                font_url = f"file://{font_path_escaped}"
            
            font_face_css = f'''
            @font-face {{
                font-family: "Arial";
                src: url("{font_url}");
            }}
            @font-face {{
                font-family: "sans-serif";
                src: url("{font_url}");
            }}
            '''
        
        # Определяем лучший шрифт для кириллицы
        # Если шрифт зарегистрирован, используем его напрямую
        if CYRILLIC_FONT_PATH:
            font_family = 'Arial, Helvetica, sans-serif'
        else:
            font_family = 'Arial, sans-serif'
        
        # Добавляем базовые стили для поддержки кириллицы, если их нет в HTML
        if '<style>' not in html_content and '<STYLE>' not in html_content:
            style_tag = f'''<style>
                {font_face_css}
                @page {{
                    size: A4;
                    margin: 2cm;
                }}
                * {{
                    font-family: {font_family} !important;
                }}
                body {{
                    font-size: 12pt;
                }}
            </style>'''
            # Вставляем стили в head, если есть, иначе в начало body
            if '</head>' in html_content or '</HEAD>' in html_content:
                html_content = html_content.replace('</head>', style_tag + '</head>').replace('</HEAD>', style_tag + '</HEAD>')
            elif '<body>' in html_content or '<BODY>' in html_content:
                html_content = html_content.replace('<body>', '<body>' + style_tag).replace('<BODY>', '<BODY>' + style_tag)
            else:
                html_content = style_tag + html_content
        else:
            # Если стили уже есть, добавляем @font-face и универсальное правило
            font_override = f'<style>{font_face_css}* {{ font-family: {font_family} !important; }}</style>'
            if '</head>' in html_content or '</HEAD>' in html_content:
                html_content = html_content.replace('</head>', font_override + '</head>').replace('</HEAD>', font_override + '</HEAD>')
            elif '<body>' in html_content or '<BODY>' in html_content:
                html_content = html_content.replace('<body>', '<body>' + font_override).replace('<BODY>', '<BODY>' + font_override)
        
        # Генерируем PDF с помощью xhtml2pdf
        # Проверяем, не открыт ли файл, и если да - удаляем старую версию
        if output_path.exists():
            try:
                output_path.unlink()
            except PermissionError:
                print(f"\n[ОШИБКА] Не удалось удалить существующий файл: {output_path}")
                print("Пожалуйста, закройте файл, если он открыт в другой программе, и попробуйте снова.")
                raise
        
        try:
            with open(output_path, 'w+b') as result_file:
                # Создаем контекст для настройки шрифтов
                # Используем link_callback для правильной обработки кодировки
                def link_callback(uri, rel):
                    return uri
                
                # Используем специальную настройку для поддержки кириллицы
                # Убеждаемся, что используется правильная кодировка
                pisa_status = pisa.CreatePDF(
                    BytesIO(html_content.encode('utf-8')),
                    dest=result_file,
                    encoding='utf-8',
                    link_callback=link_callback
                )
                
                if pisa_status.err:
                    raise Exception(f"Ошибка при создании PDF: {pisa_status.err}")
                
                # Дополнительная проверка - если есть предупреждения о шрифтах
                if hasattr(pisa_status, 'warn') and pisa_status.warn:
                    print(f"[ПРЕДУПРЕЖДЕНИЕ] При генерации PDF: {pisa_status.warn}")
            
            print(f"✓ PDF успешно создан: {output_path}")
        except PermissionError as pe:
            print(f"\n[ОШИБКА] Нет доступа к файлу: {output_path}")
            print("Возможные причины:")
            print("  1. Файл открыт в другой программе (PDF-просмотрщик, редактор и т.д.)")
            print("  2. Недостаточно прав доступа к директории")
            print("  3. Антивирус блокирует запись")
            print(f"\nПопробуйте закрыть файл {output_path.name} и запустить скрипт снова.")
            raise
    except Exception as e:
        print(f"Ошибка при генерации PDF: {e}")
        import traceback
        traceback.print_exc()
        raise


def open_pdf(file_path: Path):
    """Открывает PDF файл в системной программе."""
    try:
        import platform
        import subprocess
        
        system = platform.system()
        
        if system == 'Windows':
            os.startfile(str(file_path))
        elif system == 'Darwin':  # macOS
            subprocess.run(['open', str(file_path)])
        elif system == 'Linux':
            subprocess.run(['xdg-open', str(file_path)])
        else:
            print(f"Не удалось автоматически открыть PDF на системе {system}")
            print(f"Файл сохранен по пути: {file_path}")
    except Exception as e:
        print(f"Не удалось открыть PDF: {e}")
        print(f"Файл сохранен по пути: {file_path}")


def print_menu(title: str, items: List[Any], item_formatter=None):
    """Выводит аккуратное меню с нумерацией."""
    print(f"\n{'=' * 50}")
    print(f"  {title}")
    print(f"{'=' * 50}")
    
    if not items:
        print("  (нет доступных вариантов)")
        return
    
    for i, item in enumerate(items, 1):
        if item_formatter:
            display = item_formatter(item)
        else:
            display = str(item)
        print(f"  {i}. {display}")
    
    print(f"{'=' * 50}")


def get_user_choice(max_value: int, prompt: str = "Выберите вариант") -> Optional[int]:
    """Получает выбор пользователя с валидацией."""
    while True:
        try:
            choice = input(f"\n{prompt} (1-{max_value}): ").strip()
            if not choice:
                return None
            choice_num = int(choice)
            if 1 <= choice_num <= max_value:
                return choice_num
            else:
                print(f"Пожалуйста, введите число от 1 до {max_value}")
        except ValueError:
            print("Пожалуйста, введите корректное число")
        except KeyboardInterrupt:
            print("\n\nОперация отменена пользователем.")
            sys.exit(0)


def main():
    """Основная функция."""
    print("\n" + "=" * 50)
    print("  ГЕНЕРАТОР PDF-ДОКУМЕНТОВ")
    print("=" * 50)
    
    # Получаем списки файлов
    data_files = get_data_files()
    template_files = get_template_files()
    
    # Выводим доступные файлы данных
    print_menu(
        "ДОСТУПНЫЕ ФАЙЛЫ С ДАННЫМИ",
        data_files,
        item_formatter=lambda f: f"{f.name} ({f.suffix.upper()})"
    )
    
    if not data_files:
        print("\nОшибка: не найдено ни одного файла данных в директории /data")
        print("Пожалуйста, добавьте CSV или JSON файлы в директорию /data")
        return
    
    # Выводим доступные шаблоны
    print_menu(
        "ДОСТУПНЫЕ HTML-ШАБЛОНЫ",
        template_files,
        item_formatter=lambda f: f.name
    )
    
    if not template_files:
        print("\nОшибка: не найдено ни одного HTML-шаблона в директории /templates")
        print("Пожалуйста, добавьте HTML-шаблоны в директорию /templates")
        return
    
    # Выбор файла данных
    data_choice = get_user_choice(len(data_files), "Выберите файл с данными")
    if not data_choice:
        return
    
    selected_data_file = data_files[data_choice - 1]
    print(f"\n✓ Выбран файл данных: {selected_data_file.name}")
    
    # Загрузка данных
    print(f"Загрузка данных из {selected_data_file.name}...")
    data = load_data_file(selected_data_file)
    
    if not data:
        print("Ошибка: не удалось загрузить данные из файла")
        return
    
    print(f"✓ Загружено записей: {len(data)}")
    
    # Выбор шаблона
    template_choice = get_user_choice(len(template_files), "Выберите HTML-шаблон")
    if not template_choice:
        return
    
    selected_template = template_files[template_choice - 1]
    print(f"\n✓ Выбран шаблон: {selected_template.name}")
    
    # Получаем список invoice id
    invoice_ids = get_invoice_ids(data)
    
    if not invoice_ids:
        print("\nОшибка: не найдено ни одного invoice id в данных")
        print("Убедитесь, что в данных есть поле 'invoice_id', 'invoiceId', 'invoice', 'id' или 'ID'")
        return
    
    # Выводим список invoice id
    print_menu(
        "ДОСТУПНЫЕ ЧЕКИ (INVOICE ID)",
        invoice_ids,
        item_formatter=lambda inv_id: f"Invoice #{inv_id}"
    )
    
    # Выбор invoice id
    invoice_choice = get_user_choice(len(invoice_ids), "Выберите invoice id")
    if not invoice_choice:
        return
    
    selected_invoice_id = invoice_ids[invoice_choice - 1]
    print(f"\n✓ Выбран invoice id: {selected_invoice_id}")
    
    # Находим запись по invoice id
    record = find_record_by_invoice_id(data, selected_invoice_id)
    
    if not record:
        print("Ошибка: не удалось найти запись с выбранным invoice id")
        return
    
    # Отладочный вывод данных
    print(f"\n[DEBUG] Данные записи: {list(record.keys())}")
    print(f"[DEBUG] Значения: {record}")
    
    # Рендерим шаблон
    print(f"\nГенерация PDF для invoice #{selected_invoice_id}...")
    html_content = render_template(selected_template, record)
    
    # Генерируем PDF
    output_filename = f"invoice_{selected_invoice_id}.pdf"
    output_path = OUTPUT_DIR / output_filename
    
    try:
        generate_pdf(html_content, output_path)
        
        # Открываем PDF
        print(f"\nОткрытие PDF файла...")
        open_pdf(output_path)
        
        print(f"\n{'=' * 50}")
        print("  ГЕНЕРАЦИЯ ЗАВЕРШЕНА УСПЕШНО!")
        print(f"{'=' * 50}")
        
    except Exception as e:
        print(f"\nОшибка при генерации PDF: {e}")
        return


if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nОперация отменена пользователем.")
        sys.exit(0)
    except Exception as e:
        print(f"\nКритическая ошибка: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

