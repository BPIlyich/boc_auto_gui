## boc_auto_gui

Скрипт автоматизации формирования баз в программе BOC

## Установка:
    git clone https://github.com/PIlyichB/boc_auto_gui.git INSTALL/DIR/
    pip install -r INSTALL/DIR/requirements.txt

## Запуск:
    python INSTALL/DIR/boc_auto_gui.py PATH/TO/BOC.EXE

## Принимаемые аргументы:
+ Обязательные:
  + Путь до BOC.exe
+ Необязательные:
  + -h, --help                                  - Справка
  + -nbd BD_NAME, --bd_name BD_NAME             - Сервер БД
  + -ubd BD_USER, --bd_user BD_USER             - Пользователь БД
  + -pbd BD_PSW, --bd_psw BD_PSW                - Пароль БД
  + -eo EO                                      - Эксплуатирующая организация (по умолчанию: Ульяновский филиал ООО "Татнефть-АЗС Центр")
  + -sd START_DATE, --start_date START_DATE     - Начальная дата в формате ГГГГ-ММ-ДД (по умолчанию: Первое число текущего месяца)
  + -fd FINISH_DATE, --finish_date FINISH_DATE  - Конечная дата в формате ГГГГ-ММ-ДД (по умолчанию: Сегодняшнее число)
 
## Примечание:
Если не указаны ВСЕ аргументы из списка [Сервер БД, Пользователь БД, Пароль БД] и в BOC эти поля тоже не заполнены, скрипт завершится неудачей.
