# Split runs in Word document / Разделение ранов в Word документе

<details open>
  <summary><strong>🔷Русский🔷</strong></summary>

Этот проект представляет собой Python-скрипт для обработки документов Word (.docx) путём разбиения слов на отдельные `run` с сохранением оригинального форматирования, включая гиперссылки, сноски, рисунки и таблицы.

### Основные возможности:
- Сохранение всех стилей.
- Гиперссылки из исходного документа сохраняются и корректно включаются в итоговый документ.
- Рисунки корректно включаются в итоговый документ.
- Сноски корректно вставляются на свои места.  
- Таблицы из исходного документа переносятся без изменений.

### Требования:
- Python 3.9+  
- Библиотеки Python:  
  - `python-docx`  
  - `tqdm`  

Для установки зависимостей:  
`pip install python-docx tqdm`

## Известные прболемы

- Гиперссылки остаются такими как есть, без разделения на `run`

</details>
<details>
  <summary>🔷English🔷</summary>

This project is a Python script designed to process Word documents (.docx) by splitting words into separate `runs` while maintaining the original formatting, including hyperlinks, footnotes, figures and tables.

### Features:
- Retains all styles.
- Hyperlinks from the source document are saved and correctly included in the final document.
- Drawings are properly included in the output document.
- Footnotes are correctly inserted into their places.
- Tables from the original document are transferred without changes.

### Requirements:
- Python 3.9+  
- Python Libraries:  
  - `python-docx`  
  - `tqdm`  

To install the dependencies:  
`pip install python-docx tqdm`

## Known Issues

- Hyperlinks remain as they are, without dividing into `run`

</details>
