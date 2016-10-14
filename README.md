## boc_auto_gui

Скрипт автоматизации рутинных действий в программе BOC

## Возможности:
- Формирование и упаковка в архив MAINBASE.dbf и SVODBASE.dbf
- Формирование и упаковка в архив Stat456 и klients.dbf
- Формирование и упаковка в архив AllReports
- Формирование и упаковка в архив "Отчет для сверки с АЗС"

## Установка:
    git clone https://github.com/BPIlyich/boc_auto_gui.git INSTALL/DIR/
    pip install -r INSTALL/DIR/requirements.txt

## Запуск:
    python INSTALL/DIR/boc_auto_gui.py "PATH/TO/BOC.EXE"

## Принимаемые аргументы:
- Обязательные:
  - Путь до BOC.exe
- Необязательные:
  - -h, --help                                  - Справка
  - -v, --version                               - Версия программы

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
 
## Примечание:
- Если не указаны ВСЕ аргументы из списка [Сервер БД, Пользователь БД, Пароль БД] и в BOC эти поля тоже не заполнены, скрипт завершится неудачей.
- Тестировалось на версии BOC 8.0
