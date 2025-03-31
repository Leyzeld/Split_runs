# Split runs in Word document / Разделение ранов в Word документе

<details open>
  <summary><strong>🔷Русский🔷</strong></summary>

Этот проект представляет собой Python-скрипт для обработки документов Word (.docx) путём разбиения `run` на слова с сохранением оригинального форматирования, включая гиперссылки, сноски, рисунки и таблицы.

### Основные возможности:
- Сохранение стилей: жирный, курсив, подчёркивание, размер и цвет шрифта.
- Гиперссылки из исходного документа сохраняются.
- Рисунки корректно включаются в итоговый документ.
- Сеноски вставляются в почти правильные места 😊
- Таблицы из исходного документа переносятся без изменений.

### Требования:
- Python 3.x  
- Библиотеки Python:  
  - `python-docx`  
  - `tqdm`  

Для установки зависимостей:  
`pip install python-docx tqdm`

## Известные баги

- Заголовки не сохраняют цвет текста. Цвет текста устанавливается стилем заголовка.
- Сноски могут быть вставлены не возле соответствующего слова, но в соответствующем параграфе.
- Гиперссылки остаются такими как есть, без разделения на `run`

</details>
<details>
  <summary>🔷English🔷</summary>

This project is a Python script designed to process Word documents (.docx) by splitting `runs` into words while preserving the original formatting, including hyperlinks, footnotes, drawings, and tables.

### Features:
- Retains styles: bold, italic, underline, font size, and color.
- Hyperlinks from the original document are preserved.
- Drawings are properly included in the output document.
- Footnotes are inserted in almost correct places 😊
- Tables from the original document are transferred without changes.

### Requirements:
- Python 3.x  
- Python Libraries:  
  - `python-docx`  
  - `tqdm`  

To install the dependencies:  
`pip install python-docx tqdm`

## Known Issues

- Headings do not retain text color. The text color is defined by the heading style.
- Footnotes may not be placed next to the corresponding word, but they appear in the relevant paragraph.
- Hyperlinks remain as they are, without dividing into `run`

</details>
