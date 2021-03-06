# Change Log

## [0.3.4] - 2018-05-03
### Добавлено:
- Пустые поля NAMEEKSPL в таблице EKSLPORG.dbf заполняются автоматически исходя из эмитента (вероятно не всегда корректно)
- Аргумент -e/--emitents_path путь до ini-файла с номеров и названий эмитентов (по умолчанию: None)
### Изменено:
- Теперь жестко задано, что "Базы" и "Статистика" делаются с OpenWay, а "Отчет по клиентам" и "Отчет для сверки с АЗС" - без OpenWay
### Удалено:
- Аргумент -wow/--with_openway

## [0.3.3] - 2017-12-06
### Добавлено:
- Аргумент -wow/--with_openway - Флаг, указывающий на необходимость формирования MainBase и SvodBase с учетом OpenWay (Отчет для сверки с АЗС будет содержать транзакции OpenWay)
### Изменено:
- Теперь по умолчанию не производится вызов "Справочники->Эксплуатирующие компании" и формирование MainBase и SvodBase. Эти действия будут выполнены при наличии любого из следующих флагов: zip_mainbase_svodbase / zip_stat / zip_client_reports / zip_gs_report

## [0.3.2] - 2017-03-03
### Исправленно:
- Проблема с относительными путями в конфигурационном файле
### Изменено:
- Откатил версию pywinauto до 0.5.4, т.к. на Windows XP появляются ошибки (из-за отсутствия uia?)


## [0.3.1] - 2017-02-22
### Добавлено:
- Аргумент -npd/--no_period_dialog - Флаг, убирающий диалог выбора дат при запуске
### Изменено:
- Теперь по умолчанию требуется ввод/проверка дат через диалог, если не указан флаг --no_period_dialog


## [0.3.0] - 2017-02-22
### Добавлено:
- Собрал все в один exe (с помощью PyInstaller)
- Аргумент -с/--config для считывания настроек из файла (формат файла - INI с кодировкой cp1251)
- Аргумент -mh/--mail_host - Хост электронной почты (по умолчанию: None)
- Аргумент -mp/--mail_port - Порт для почты (по умолчанию: None)
- Аргумент -fa/--from_addr - Почта с которой производим рассылку (по умолчанию: None)
- Аргумент -mpsw/--mail_password - Пароль от почты (по умолчанию: None)
- Аргумент -ba/--bcc_addr - Почта для отправки "слепой" копии (по умолчанию: None)
- Аргумент -ra/--reply_addr - Почта для ответа (по умолчанию: None)
- Аргумент -ab/--to_addr_bases - Почта для отправки базы (по умолчанию: None)
- Аргумент -ac/--to_addr_client_reports - Почта для отправки отчета по клиентам (по умолчанию: None)
- Аргумент -ag/--to_addr_gs_report - Почта для отправки отчета по АЗС (по умолчанию: None)
- Аргумент -as/--to_addr_stat - Почта для отправки статистики (по умолчанию: None)
### Изменено:
- Параметр -eo теперь имеет и полную форму --ekspl_org (чтобы можно было задавать параметр через файл конфигурации)
- Единственный позиционный аргумент boc_path стал именованным (-b/--boc_path) (чтобы можно было задавать параметр через файл конфигурации)


## [0.2.2] - 2016-11-30
### Изменено:
- Отчеты и скриншоты ошибок теперь попадают в папку output

## [0.2.1] - 2016-10-14
### Исправлено:
- Ошибка поиска окна "Статистика"

### Изменено:
- Даты по умолчанию:
    - начальная - первый день предыдущего месяца
    - конечная - первый день текущего месяца

## [0.2.0] - 2016-10-07
### Добавлено:
- CHANGELOG.md
- Возможность паковать в архив MAINBASE.dbf и SVODBASE.dbf
- Возможность формировать и паковать в архив Stat456 и klients.dbf
- Возможность формировать и паковать в архив AllReports
- Возможность формировать и паковать в архив "Отчет для сверки с АЗС"

## [0.1.0] - 2016-09-29
### Добавлено:
- Возможность формировать MAINBASE.dbf и SVODBASE.dbf