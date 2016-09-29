# -*- coding: utf-8 -*-

__version__ = '0.1'

import argparse
import datetime
import logging
import os
from pywinauto.application import Application, AppStartError
from pywinauto.timings import WaitUntil
import dbf


################################### logging ####################################
# у pywinauto есть свои настройки логгирования по умолчанию для консоли
# поэтому добавим только логгирование в файл
LOG_FILE = os.path.join(
    os.path.dirname(os.path.realpath(__file__)),
    'boc_auto.log'
)
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
log_formatter = logging.Formatter('[%(asctime)s] [%(levelname)-8s] %(message)s')
filehandler = logging.FileHandler(LOG_FILE)
filehandler.setFormatter(log_formatter)
logger.addHandler(filehandler)
################################################################################


def get_dbf_empty_field_value(table, fieldname):
    field_length = table.field_info(fieldname)[1]
    return ' ' * field_length


def fill_eo_in_dbf(table, eo):
    count = 0
    logger.info(u'Открываем klients.dbf для заполнения ЭО')
    with table:
        ind = table.create_index(lambda rec: rec.nameekspl)
        empty_value = get_dbf_empty_field_value(table, 'nameekspl')
        records = ind.search(match=(empty_value, ))
        for record in dbf.Process(records):
            logger.info(u'Заполняем ЭО для клиента {} с id {}'.format(
                record.name, record.id_klienta))
            record.nameekspl = eo
            count += 1
    return count


def valid_file(path):
    if os.path.isfile(path):
        return path
    raise argparse.ArgumentTypeError('File {} doesnot exists'.format(path))


def valid_date(s):
    try:
        return datetime.datetime.strptime(s, '%Y-%m-%d').date()
    except ValueError:
        msg = 'Invalid date: "{}".'.format(s)
        raise argparse.ArgumentTypeError(msg)


def auto_gui(boc_path, bd_name, bd_user, bd_psw, eo, start_date, finish_date):
    try:
        working_dir = os.path.dirname(boc_path)
        os.chdir(working_dir)

        logger.info(u'Запускаем программу')
        app = Application().Start(cmd_line=boc_path)
        menu = app['TOraTypeMain']
        menu.Wait('ready')

        logger.info(u'Подключаемся к серверу')
        menu.MenuSelect(u'Файл->Подключиться к серверу')
        connect_to_bd_window = app[u'Подключение к БД']
        if all([bd_name, bd_user, bd_psw]):
            connect_to_bd_window['Edit3'].SetText(bd_name)
            connect_to_bd_window['Edit2'].SetText(bd_user)
            connect_to_bd_window['Edit1'].SetText(bd_psw)
        connect_to_bd_window[u'ОК'].Click()
        #connect_to_bd_window.WaitNot('visible')
        menu.Wait('ready')

        logger.info(u'Заполняем Справочник Клиентов')
        menu.MenuSelect(u'Справочники->Справочник Клиентов')
        clients_window = app[u'Справочник клиентов.']
        clients_window.Wait('exists')
        clients_window[u'Добавить'].Click()
        clients_window['TDBNavigator'].Wait('ready')
        clients_window[u'ОК'].SetFocus()
        clients_window[u'ОК'].Click()
        clients_window.WaitNot('visible')
        menu.Wait('ready')

        table_path = os.path.join(working_dir, 'klients.dbf')
        table = dbf.Table(table_path, codepage='cp1251')
        count = fill_eo_in_dbf(table, eo)
        if count:
            logger.info(u'Заполнили столбец "Чей клиент" для {} строк'.format(
                count))
        else:
            logger.info(u'Столбец "Чей клиент" заполнять не требуется')

        logger.info(u'Заполняем Эксплуатирующие компании')
        menu.MenuSelect(u'Справочники->Эксплуатирующие компании')
        eo_window = app[u'Эксплуатирующие компании']
        eo_window.Wait('exists')
        eo_window[u'Дополнить'].Click()
        # XXX тут надо подождать?
        eo_window[u'Закрыть'].Click()
        eo_window.WaitNot('visible')
        menu.Wait('ready')

        def is_task_finished():
            try:
                if menu['StatusBar'].Exists():
                    return menu['StatusBar'].WindowText().strip().startswith(
                        u'Затрачено времени')
            except TypeError:
                pass
            return False

        logger.info(u'Переносим справочники в Oracle')
        menu.MenuSelect(u'Новые запросы->Перенос справочников в &Oracle')
        WaitUntil(180, 5, is_task_finished)

        logger.info(
            u'Формируем MAINBASE.dbf и SVODBASE.dbf '
            u'с {:%d.%m.%Y} по {:%d.%m.%Y}'.format(start_date, finish_date))
        menu.MenuSelect(
            u'Новые запросы->Сформировать M&AINBASE.dbf и SVODBASE.dbf')
        date_period_window = app[u'Период']
        date_period_window.Wait('exists')
        date_period_window['TDateTimePicker2'].SetFocus()
        date_period_window['TDateTimePicker2'].TypeKeys(
            '{:%d{{RIGHT}}%m{{RIGHT}}%Y}'.format(start_date))
        date_period_window['TDateTimePicker1'].SetFocus()
        date_period_window['TDateTimePicker1'].TypeKeys(
            '{:%d{{RIGHT}}%m{{RIGHT}}%Y}'.format(finish_date))
        date_period_window[u'ОК'].Click()
        WaitUntil(300, 5, is_task_finished)

        logger.info(u'Закрываем BOC')
        menu.MenuSelect(u'Файл->Выход')
        logger.info(u'Все операции успешно завершены')
    except AppStartError:
        logger.exception(u'Не удалось запустить приложение')
        raise
    except:
        logger.exception(u'Неожиданное завершение программы')
        img_path = os.path.join(
            os.path.dirname(LOG_FILE),
            'error_{:%Y%m%d_%H%M%S}.png'.format(datetime.datetime.now())
        )
        try:
            app.top_window_().CaptureAsImage().save(img_path)
            logger.info(u'Скриншот с ошибкой - {}'.format(img_path))
        except AttributeError:
            logger.warning(u'Похоже не установлен PIL/Pillow. '
                           u'Невозможно сделать скриншот')
        except RuntimeError:
            logger.exception(u'Скриншот программы сделать не удалось')
        # Если таблица EKSPLORG.dbf занята другим процессом,
        # BOC невозможно завершить корректно!
        logger.warning(u'Убиваем процесс...')
        app.Kill_()
        raise


if __name__ == '__main__':
    today = datetime.date.today()
    first_day = datetime.date(today.year, today.month, 1)

    parser = argparse.ArgumentParser()
    parser.add_argument('boc_path', type=valid_file, help=u'Путь до BOC')
    parser.add_argument('-nbd', '--bd_name', help=u'Сервер БД')
    parser.add_argument('-ubd', '--bd_user', help=u'Пользователь БД')
    parser.add_argument('-pbd', '--bd_psw', help=u'Пароль БД')
    parser.add_argument(
        '-eo', help=u'Эксплуатирующая организация (по умолчанию:  %(default)s)',
        default=u'Ульяновский филиал ООО "Татнефть-АЗС Центр"'
    )
    parser.add_argument(
        '-sd', '--start_date', type=valid_date, default=first_day,
        help=u'Начальная дата в формате ГГГГ-ММ-ДД (по умолчанию: %(default)s)'
    )
    parser.add_argument(
        '-fd', '--finish_date', type=valid_date, default=today,
        help=u'Конечная дата в формате ГГГГ-ММ-ДД (по умолчанию: %(default)s)'
    )

    args = parser.parse_args()
    logger.info(u'Переданные аргументы: {}'.format(args))
    auto_gui(**vars(args))
