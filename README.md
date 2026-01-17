![Python](https://img.shields.io/badge/Python-3.9+-blue?logo=python&logoColor=white)
![PDF](https://img.shields.io/badge/PDF-Generator-red)
![Automation](https://img.shields.io/badge/Automation-Enabled-orange)
![xhtml2pdf](https://img.shields.io/badge/xhtml2pdf-Library-green)
![Status](https://img.shields.io/badge/Status-Active-success)

# Генератор PDF-документов

Скрипт для генерации PDF-документов из CSV/JSON данных и HTML-шаблонов.

## Установка

1. Установите зависимости:
```bash
pip install -r requirements.txt
```

**Примечание:** Скрипт использует библиотеку `xhtml2pdf`, которая не требует системных зависимостей и работает на Windows, macOS и Linux без дополнительной настройки.

### Поддержка кириллицы

Скрипт автоматически пытается зарегистрировать шрифт Arial для поддержки кириллицы. Если кириллица отображается как черные квадраты:

1. **На Windows:** Убедитесь, что шрифт Arial установлен (обычно установлен по умолчанию)
2. **На macOS:** Установите шрифт Arial или используйте системные шрифты
3. **На Linux:** Установите шрифты с поддержкой кириллицы:
   ```bash
   sudo apt-get install fonts-dejavu fonts-liberation
   ```

Если проблема сохраняется, попробуйте:
- Использовать другой шаблон
- Проверить, что данные в CSV/JSON содержат корректную кодировку UTF-8

## Структура проекта

```
.
├── pdf_generator.py    # Основной скрипт
├── requirements.txt    # Зависимости Python
├── data/              # CSV и JSON файлы с данными
├── templates/         # HTML-шаблоны
└── output/           # Сгенерированные PDF файлы
```

## Использование

1. Поместите файлы с данными (CSV или JSON) в директорию `data/`
2. Поместите HTML-шаблоны в директорию `templates/`
3. Запустите скрипт:
```bash
python pdf_generator.py
```

4. Следуйте инструкциям в консоли:
   - Выберите файл с данными
   - Выберите HTML-шаблон
   - Выберите invoice id
   - PDF будет автоматически сгенерирован и открыт

## Формат данных

### CSV файлы
Должны содержать колонку с invoice id (может называться: `invoice_id`, `invoiceId`, `invoice`, `id`, `ID`)

### JSON файлы
Могут быть:
- Массивом объектов: `[{...}, {...}]`
- Одиночным объектом: `{...}`

Каждый объект должен содержать поле с invoice id.

## HTML-шаблоны

Используйте стандартный HTML с подстановкой данных через фигурные скобки:
```html
<h1>Invoice #{invoice_id}</h1>
<p>Клиент: {customer_name}</p>
<p>Сумма: {amount}</p>
```

Для экранирования фигурных скобок используйте двойные: `{{` и `}}`

## Примеры

См. файлы в директориях `data/` и `templates/` для примеров формата данных и шаблонов.


