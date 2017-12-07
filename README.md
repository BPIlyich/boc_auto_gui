## boc_auto_gui

Скрипт автоматизации рутинных действий в программе BOC

## Возможности:
- Формирование и упаковка в архив MAINBASE.dbf и SVODBASE.dbf
- Формирование и упаковка в архив Stat456 и klients.dbf
- Формирование и упаковка в архив AllReports
- Формирование и упаковка в архив "Отчет для сверки с АЗС"
- Заполнение ЭО клиентов и перенос справочников в Oracle

## Установка:
    git clone https://github.com/BPIlyich/boc_auto_gui.git INSTALL/DIR/
    pip install -r INSTALL/DIR/requirements.txt

## Сборка exe из исходников:
Сборка exe для последней версии производится с помощью [pyinstaller 3.2.1](https://pypi.python.org/pypi/PyInstaller/3.2.1), т.к. exe, собранный с помощью pyinstaller 3.3, не работает в Windows XP.

    pyinstaller INSTALL/DIR/boc_auto_gui.py --distpath OUTPUT/DIR/FOR/EXE --onefile

## Запуск:
    python INSTALL/DIR/boc_auto_gui.py -b "PATH/TO/BOC.EXE"

## Принимаемые аргументы:
- Обязательные:
  - -b BOC_PATH, --boc_path BOC_PATH            - Путь до BOC.exe
- Необязательные:
  - -h, --help                                  - Справка
  - -v, --version                               - Версия программы

  - -c CONFIG, --config CONFIG                  - Путь до файла настроек
  - -nbd BD_NAME, --bd_name BD_NAME             - Сервер БД
  - -ubd BD_USER, --bd_user BD_USER             - Пользователь БД
  - -pbd BD_PSW, --bd_psw BD_PSW                - Пароль БД
  - -eo EO                                      - Эксплуатирующая организация (по умолчанию: Ульяновский филиал ООО "Татнефть-АЗС Центр")
  - -sd START_DATE, --start_date START_DATE     - Начальная дата в формате ГГГГ-ММ-ДД (по умолчанию: Первое число предыдущего месяца)
  - -fd FINISH_DATE, --finish_date FINISH_DATE  - Конечная дата в формате ГГГГ-ММ-ДД (по умолчанию: Первое число текущего месяца)
  - -zms, --zip_mainbase_svodbase               - Запаковать в архив MAINBASE.dbf и SVODBASE.dbf
  - -zst, --zip_stat                            - Запаковать в архив Stat456 и klients.dbf
  - -zcr, --zip_client_reports                  - Запаковать в архив AllReports
  - -zgr, --zip_gs_report                       - Запаковать в архив "Отчет для сверки с АЗС"
  - -wow, --with_openway                        - Формирование MainBase и SvodBase с учетом OpenWay (на данный момент актуально только вместе с флагом --zip_gs_report)

## Примечание:
- Если не указаны ВСЕ аргументы из списка [Сервер БД, Пользователь БД, Пароль БД] и в BOC эти поля тоже не заполнены, скрипт завершится неудачей.
- Тестировалось на версии BOC 8.0
- Консольные аргументы имеют приоритет над параметрами из файла настроек
- Файл настроек должен быть в кодировке cp1251