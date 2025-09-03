# jira_workload_converter

Скрипт предназначен для конвертирования выгрузки данных из Jira в формат "нагрузки"

## Зависимости
* [Python 3.11.9](https://www.python.org/downloads/release/python-3119/)
* Microsoft Excel 2007 или выше
* Библиотеки из [requirements.txt](requirements.txt)

## Подготовка и запуск
### Установка Python
Python нужной версии можно скачать на [официальном сайте](https://www.python.org/downloads/release/python-3119/).

Если python нужной версии уже установлен на компьютере в этом можно убедиться запустив
```powershell
py --version
```
или посмотреть все установленные версии Python
```powershell
py --list
```

Шаги установки:
* Первый шаг
  * Поставить галочку "Add python.exe to PATH"
  * Поставить галочку "Use admin privileges when installing py.exe" (не нужно если уже установлен)
  * Нажать "Customize installation"
* Второй шаг
  * Обязательно поставить галочки (остальные по желанию, на работу скрипта не должно повлиять) у:
    * pip
    * py launcher (если еще не установлен)
  * Нажать Next
* Третий шаг
  * Обязательно поставить галочки (остальные по желанию, на работу скрипта не должно повлиять) у:
    * Add Python to environment variables
  * Нажать Install
* После успешной уставки закрыть окно

После установки, вероятно, потребуется перезапустить компьютер (из-за изменения переменной PATH)

Проверить успешную установку можно через powershell или cmd:
```powershell
py --version
```

Вызвать powershell или cmd можно с помощью <kbd>WIN</kbd>+<kbd>R</kbd>, вписав туда "powershell" или "cmd" соответственно.

### Установка зависимостей (библиотек)
Для установки библиотек можно использовать менеджер пакетов PIP (обычно устанавливается вместе с Python).

Находясь в директории проекта выполнить в консоли:
```powershell
py -m pip install -r requirements.txt
```

Чтобы запустить powershell в конкретной директории можно нажать <kbd>SHIFT</kbd>+<kbd>ПКМ</kbd> >> Открыть окно powershell здесь.

### Запуск скрипта

#### Описание (help) скрипта:
```powershell
usage: jira_converter.py [-h] [-o OUTPUT_FILENAME] [-y YEAR] [--freeze-cell FREEZE_CELL] filename

positional arguments:
  filename              Input filename (must be xlsx)

options:
  -h, --help            show this help message and exit
  -o OUTPUT_FILENAME, --output-filename OUTPUT_FILENAME
                        Output filename (must be xlsx)
  -y YEAR, --year YEAR  Year for data generation
  --freeze-cell FREEZE_CELL
                        Cell cords for freezing rows above and columns to the left (ex: "3,3" or "12, 34")
```

#### Подробное описание параметров
* `filename` (обязательный, порядковый) - Имя файла выгрузки из Jira. Обязательно должен иметь формат xlsx.
* `-h` или `--help` (не обязательный) - Вывести help скрипта и завершить работу.
* `-o` или `--output-filename` (не обязательный) - Имя файла, в который записать результат обработки. Обязательно должен иметь формат xlsx. Стандартное значение: "output.xlsx".
* `-y` или `--year` (не обязательный) - Год, для которого требуется обработать данные. Должен быть положительным числом. Стандартное значение: текущий год.
* `--freeze-cell` (не обязательный) - Ячейка, слева и сверху от которой закрепятся строки и столбцы (для заголовков). Стандартное значение: (4,4), то есть D4.
