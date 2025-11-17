# Word to PowerPoint Converter

Автоматично конвертира данни от Word документи в PowerPoint презентации.

## Възможности

- Извлича данни от Word документ (номер на проект, дати, цена и др.)
- Попълва PowerPoint бланка с извлечените данни
- Автоматично именува изходния файл по номера на проекта
- Лесен за употреба графичен интерфейс
- Поддържа и команден ред

## Инсталация

### 1. Инсталирай Python

Трябва да имаш инсталиран Python 3.7 или по-нова версия.

- Свали от: https://www.python.org/downloads/
- При инсталация избери опцията "Add Python to PATH"

### 2. Инсталирай необходимите библиотеки

Отвори Command Prompt (CMD) или Terminal и изпълни:

```bash
pip install -r requirements.txt
```

Или инсталирай ръчно:

```bash
pip install python-docx python-pptx Pillow lxml
```

## Употреба

### Графичен интерфейс (препоръчително)

1. Двоен клик на `word_to_ppt_gui.py` или изпълни:

```bash
python word_to_ppt_gui.py
```

2. Избери Word документа който искаш да конвертираш
3. (Опционално) Избери PowerPoint бланка (по подразбиране използва "Project 742_051.pptx")
4. (Опционално) Избери папка където да се запише резултата
5. Кликни "Конвертирай"

### Команден ред

```bash
python word_to_ppt_converter.py <word_file> [ppt_template] [output_dir]
```

Примери:

```bash
# Основна употреба (използва стандартната бланка)
python word_to_ppt_converter.py 02_785_692_OFFER_APPROVED.docx

# С конкретна бланка
python word_to_ppt_converter.py 02_785_692_OFFER_APPROVED.docx "Project 742_051.pptx"

# Със специфична изходна папка (например D:\Projects)
python word_to_ppt_converter.py 02_785_692_OFFER_APPROVED.docx "Project 742_051.pptx" "D:\Projects"
```

## Настройка на бланката

За да работи автоматизацията:

1. Запази PowerPoint бланката в папката с програмата (или на D:)
2. Бланката трябва да съдържа:
   - Заглавие с "Project XXX_XXX" (където XXX_XXX е номер на проект)
   - Finance таблица с "Required Capex", "Required Opex", "TOTAL"
   - Initial Planning секция с дати

## Какви данни се извличат от Word?

Програмата автоматично извлича:

- **Номер на проект** (например: 785_692)
- **Approval date** (дата на одобрение)
- **Finish of the project (estimated)** (приблизителна крайна дата)
- **Total cost** (обща цена)
- Допълнителни полета: Plant owner, Plant code, Name of the part, Reference, Tool number, SE inventory number, Type of service

## Какво се попълва в PowerPoint?

Програмата автоматично попълва:

- **Заглавие** - заменя номера на проекта (например: "Project 742_051" → "Project 785_692")
- **Finance таблица** - попълва Required Capex, Required Opex, TOTAL със стойността от Total cost
- **Initial Planning** - добавя датите за одобрение и крайна дата

## Изходен файл

Създаденият PowerPoint файл автоматично се именува като:

```
Project <номер_на_проект>.pptx
```

Например: `Project 785_692.pptx`

## Проблеми и решения

### Грешка: "Module not found"

Решение:
```bash
pip install python-docx python-pptx
```

### Грешка: "Template not found"

Решение: Провери че PowerPoint бланката е в същата папка като програмата, или укажи пълния път.

### Word документът е празен

Решение: Уверете се че Word файлът е попълнен с данни, а не само шаблон.

## Структура на файловете

```
├── word_to_ppt_gui.py           # Графичен интерфейс
├── word_to_ppt_converter.py     # Основна логика
├── requirements.txt             # Необходими библиотеки
├── Project 742_051.pptx         # PowerPoint бланка (трябва да бъде в папката)
└── README.md                    # Тази документация
```

## Допълнителни инструкции

### Поставяне на бланката на D:

Ако искаш бланката да е на D: диск вместо в папката на програмата:

1. Копирай "Project 742_051.pptx" на D:\
2. При стартиране на GUI апликацията, натисни "Промени бланка" и избери бланката от D:\

## Технически детайли

Програмата е написана на Python 3 и използва:
- `python-docx` - за четене на Word документи
- `python-pptx` - за редактиране на PowerPoint файлове
- `tkinter` - за графичния интерфейс

## Автор

Създадено с Claude Code
