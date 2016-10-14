# -*- coding: utf-8 -*-

__version__ = '0.2.1'

import argparse
import datetime
import logging
import os
import zipfile

import dbf
from pywinauto.application import Application, AppStartError
from pywinauto.timings import WaitUntil


NOW = datetime.datetime.now()
FILE_DIR_PATH = os.path.dirname(os.path.realpath(__file__))


################################### logging ####################################
# у pywinauto есть свои настройки логгирования по умолчанию для консоли
# поэтому добавим только логгирование в файл
LOG_FILE = os.path.join(FILE_DIR_PATH, 'boc_auto.log')
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
log_formatter = logging.Formatter('[%(asctime)s] [%(levelname)-8s] %(message)s')
filehandler = logging.FileHandler(LOG_FILE)
filehandler.setFormatter(log_formatter)
logger.addHandler(filehandler)

logging.basicConfig(
    format='[%(asctime)s] [%(levelname)-8s] %(message)s',
    level=logging.INFO)
################################################################################


def compress_to_zip(input_files, output_file):
    if not isinstance(input_files, (list, tuple)):
        input_files = [input_files]
    with zipfile.ZipFile(output_file, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for path in input_files:
            if os.path.isdir(path):
                for root, dirs, files in os.walk(path):
                    for f in files:
                        relpath = os.path.relpath(
                            os.path.join(root, f), os.path.dirname(path))
                        zip_file.write(os.path.join(root, f), relpath)
            elif os.path.isfile(path):
                relpath = os.path.basename(path)
                zip_file.write(path, relpath)
            else:
                logger.warning('{} - Путь не найден'.format(path))


def get_dbf_empty_field_value(table, fieldname):
    field_length = table.field_info(fieldname)[1]
    return ' ' * field_length


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


class BocAutoGui():
    def __init__(self, boc_path, bd_name, bd_user, bd_psw, eo,
                 start_date, finish_date, zip_mainbase_svodbase, zip_stat,
                 zip_client_reports, zip_gs_report):
        self.boc_dir_path = os.path.dirname(boc_path)
        self.bd_name = bd_name
        self.bd_user = bd_user
        self.bd_psw = bd_psw
        self.eo = eo
        self.start_date = start_date
        self.finish_date = finish_date
        self.need_archive_mainbase_and_svodbase = zip_mainbase_svodbase
        self.need_client_reports = zip_client_reports
        self.need_gs_report = zip_gs_report
        self.need_stat = zip_stat
        try:
            self.app = Application().Start(cmd_line=boc_path)
        except AppStartError:
            logger.exception(u'Не удалось запустить приложение')
            raise
        self.menu = self.app['TOraTypeMain']

    def kill_process(self):
        logger.warning(u'Убиваем процесс...')
        self.app.Kill_()

    def make_top_window_screenshot(self):
        img_path = os.path.join(
            FILE_DIR_PATH,
            'error_{:%Y%m%d_%H%M%S}.png'.format(datetime.datetime.now())
        )
        try:
            self.app.top_window_().CaptureAsImage().save(img_path)
            logger.info(u'Скриншот с ошибкой - {}'.format(img_path))
        except AttributeError:
            logger.warning(u'Похоже не установлен PIL/Pillow. '
                           u'Невозможно сделать скриншот')
        except:
            logger.exception(u'Скриншот программы сделать не удалось')

    def fill_eo_in_klients_dbf(self):
        count = 0
        table_path = os.path.join(self.boc_dir_path, 'klients.dbf')
        table = dbf.Table(table_path, codepage='cp1251')
        logger.info(u'Открываем klients.dbf для заполнения ЭО')
        with table:
            ind = table.create_index(lambda rec: rec.nameekspl)
            empty_value = get_dbf_empty_field_value(table, 'nameekspl')
            records = ind.search(match=(empty_value, ))
            for record in dbf.Process(records):
                logger.info(u'Заполняем ЭО для клиента {} с id {}'.format(
                    record.name, record.id_klienta))
                record.nameekspl = self.eo
                count += 1
        if count:
            logger.info(u'Заполнили столбец "Чей клиент" для {} строк'.format(
                count))
        else:
            logger.info(u'Столбец "Чей клиент" заполнять не требуется')

    def fix_mainbase_dbf(self):
        table_path = os.path.join(self.boc_dir_path, 'MAINBASE.dbf')
        table = dbf.Table(table_path, codepage='cp1251')
        logger.info(u'Открываем MAINBASE.dbf для исправления проблемы с датами')
        with table:
            for record in dbf.Process(table):
                record.vdata = record.data

    def is_task_finished(self):
        try:
            if self.menu['StatusBar'].Exists():
                return self.menu['StatusBar'].WindowText().strip().startswith(
                    u'Затрачено времени')
        except TypeError:
            pass
        return False

    def is_status_ok(self):
        try:
            if self.menu['StatusBar'].Exists():
                return self.menu['StatusBar'].WindowText().strip().startswith(
                    u'ОК!')
        except TypeError:
            pass
        return False

    def fill_date_period(self):
        date_period_window = self.app[u'Период']
        date_period_window.Wait('exists')
        date_period_window['TDateTimePicker2'].SetFocus()
        date_period_window['TDateTimePicker2'].TypeKeys(
            '{:%d{{RIGHT}}%m{{RIGHT}}%Y}'.format(self.start_date))
        date_period_window['TDateTimePicker1'].SetFocus()
        date_period_window['TDateTimePicker1'].TypeKeys(
            '{:%d{{RIGHT}}%m{{RIGHT}}%Y}'.format(self.finish_date))
        date_period_window[u'ОК'].Click()

    def login(self):
        logger.info(u'Подключаемся к серверу')
        self.menu.MenuSelect(u'Файл->Подключиться к серверу')
        connect_to_bd_window = self.app[u'Подключение к БД']
        if all([self.bd_name, self.bd_user, self.bd_psw]):
            connect_to_bd_window['Edit3'].SetText(self.bd_name)
            connect_to_bd_window['Edit2'].SetText(self.bd_user)
            connect_to_bd_window['Edit1'].SetText(self.bd_psw)
        connect_to_bd_window[u'ОК'].Click()
        self.menu.Wait('ready')

    def fill_clients(self):
        logger.info(u'Дополняем Справочник Клиентов')
        self.menu.MenuSelect(u'Справочники->Справочник Клиентов')
        clients_window = self.app[u'Справочник клиентов.']
        clients_window.Wait('exists')
        clients_window[u'Добавить'].Click()
        clients_window['TDBNavigator'].Wait('ready')
        clients_window[u'ОК'].SetFocus()
        clients_window[u'ОК'].Click()
        clients_window.WaitNot('visible')
        self.menu.Wait('ready')

        self.fill_eo_in_klients_dbf()

    def fill_companies(self):
        logger.info(u'Дополняем Эксплуатирующие компании')
        self.menu.MenuSelect(u'Справочники->Эксплуатирующие компании')
        eo_window = self.app[u'Эксплуатирующие компании']
        eo_window.Wait('exists')
        eo_window[u'Дополнить'].Click()
        # XXX timeout?
        eo_window[u'Закрыть'].Click()
        eo_window.WaitNot('visible')
        self.menu.Wait('ready')

    def fill_oracle(self):
        logger.info(u'Переносим справочники в Oracle')
        self.menu.MenuSelect(u'Новые запросы->Перенос справочников в &Oracle')
        WaitUntil(180, 5, self.is_task_finished)

    def form_mainbase_and_svodbase(self):
        logger.info(
            u'Формируем MAINBASE.dbf и SVODBASE.dbf '
            u'с {:%d.%m.%Y} по {:%d.%m.%Y}'.format(
                self.start_date, self.finish_date)
        )
        self.menu.MenuSelect(
            u'Новые запросы->Сформировать M&AINBASE.dbf и SVODBASE.dbf')
        self.fill_date_period()
        WaitUntil(300, 5, self.is_task_finished)

    def form_mainlnr(self):
        logger.info(u'Формируем MAIN_LNR.dbf')
        self.menu.MenuSelect(u'Новые запросы->Сформировать &MAIN_LNR.dbf')
        self.fill_date_period()
        WaitUntil(300, 5, self.is_task_finished)

    def calc_stat(self):
        logger.info(u'Запускаем Расчет статистики')
        self.menu.MenuSelect(u'Подготовить данные->Расчет статистики')
        #stat_window = self.app[u'Статистика']  # <-- Иногда с другим названием?
        stat_window = self.app.TFrmStatistics
        stat_window.Wait('exists')
        stat_window[u'Расчет статистических показателей'].Click()
        #stat_window[u'Закрыть'].Wait('enabled')  # <-- Не работает

        def is_ok():
            try:
                return stat_window['StatusBar'].WindowText().strip().startswith(
                    u'Расчет показателей завершен !')
            except TypeError:
                pass
            return False
        WaitUntil(300, 5, is_ok)
        stat_window[u'Закрыть'].Click()
        self.menu.Wait('ready')

    def make_client_reports(self):
        logger.info(u'Формируем отчёты для всех клиентов без ID')
        self.menu.MenuSelect(
            u'Отчеты->Отчеты по Клиенту->'
            u'Подготовить файлы отчетов для всех клиентов без I&D'
        )
        clients_window2 = self.app[u'Параметры']
        clients_window2.Wait('exists')
        clients_window2['ComboBox'].Select(self.eo)
        clients_window2[u'Фильтр'].Click()
        clients_window2[u'Выбрать все'].Click()
        # XXX timeout?
        clients_window2[u'Ок'].Click()

        checkbox_window = self.app[u'Выбор отчетов для печати']
        checkbox_window.Wait('exists')
        checkbox_window[u'Отчет "По Клиенту и Цене" (общий)'].CheckByClick()
        checkbox_window[
            u'Отчет "По Клиенту с расшифровкой" (общий)'].CheckByClick()
        checkbox_window[
            u'Отчет "По Клиенту без расшифровки"  (общий)'].UncheckByClick()
        checkbox_window[u'Отчет "По Клиенту и Региону"'].UncheckByClick()
        checkbox_window[u'Отчет по предоставленным скидкам'].UncheckByClick()
        checkbox_window[u'Формировать отчеты в EXCEL'].UncheckByClick()
        checkbox_window[u'OK'].Click()
        WaitUntil(300, 5, self.is_status_ok)

    def make_gs_report(self):
        self.fix_mainbase_dbf()

        # XXX только для своих, т.к. не умеем читать значения из TBGRID
        logger.info(u'Формируем Отчет для сверки с АЗС (Приложение 8)')
        self.menu.MenuSelect(
            u'Отчеты->Отчеты для ЭО->Отчет для сверки с АЗС (Приложение &8)')
        gs_window = self.app[u'Параметры']
        gs_window.Wait('exists')
        gs_window['ComboBox'].Select(self.eo)
        gs_window[u'Отбор'].Click()
        # XXX timeout?
        gs_window[u'Выбрать свои'].Click()
        gs_window[u'Ок'].Click()
        gs_window.WaitNot('visible')

        output_pathes = []
        params_window = self.app[u'Параметры']
        preview_window = self.app[u'Preview']
        save_window = self.app[u'Сохранить как']
        for i in range(2):  # 2 диалога подряд
            filename = (
                'gs_report-{:%Y%m%d}-{:%Y%m%d}_{:%Y%m%d%H%M%S}-{}.frp'.format(
                    self.start_date, self.finish_date, NOW, i+1
                )
            )
            output_path = os.path.join(FILE_DIR_PATH, filename)
            params_window.Wait('exists')
            params_window[u'Ок'].Click()

            preview_window.Wait('exists')
            preview_window.TypeKeys('^s')  # Вызываем диалог сохранения в файл

            save_window.Wait('exists')
            save_window['Edit'].SetEditText(output_path)
            save_window[u'Со&хранить'].Click()

            preview_window.Close()
            output_pathes.append(output_path)
        return output_pathes

    def make_base_zip(self):
        input_files = (
            os.path.join(self.boc_dir_path, 'MAINBASE.dbf'),
            os.path.join(self.boc_dir_path, 'SVODBASE.dbf'),
        )
        output_file = os.path.join(
            FILE_DIR_PATH,
            'base456_{:%Y%m%d}-{:%Y%m%d}_{:%Y%m%d%H%M%S}.zip'.format(
                self.start_date, self.finish_date, NOW
            )
        )
        logger.info(u'Пакуем файлы MAINBASE.dbf и SVODBASE.dbf '
                    u'в архив - {}'.format(output_file))
        compress_to_zip(input_files, output_file)

    def make_stat_zip(self):
        input_files = (
            os.path.join(self.boc_dir_path, 'Stat456'),
            os.path.join(self.boc_dir_path, 'klients.dbf'),
        )
        output_file = os.path.join(
            FILE_DIR_PATH,
            'stat456_{:%Y%m%d}-{:%Y%m%d}_{:%Y%m%d%H%M%S}.zip'.format(
                self.start_date, self.finish_date, NOW
            )
        )
        logger.info(u'Пакуем файлы Stat456 и klients.dbf '
                    u'в архив - {}'.format(output_file))
        compress_to_zip(input_files, output_file)

    def make_client_reports_zip(self):
        input_files = os.path.join(self.boc_dir_path, 'AllReports')
        output_file = os.path.join(
            FILE_DIR_PATH,
            'reports_{:%Y%m%d}-{:%Y%m%d}_{:%Y%m%d%H%M%S}.zip'.format(
                self.start_date, self.finish_date, NOW
            )
        )
        logger.info(u'Пакуем файлы AllReports в архив - {}'.format(
            output_file))
        compress_to_zip(input_files, output_file)

    def make_gs_report_zip(self, input_files):
        output_file = os.path.join(
            FILE_DIR_PATH,
            'gs_report_{:%Y%m%d}-{:%Y%m%d}_{:%Y%m%d%H%M%S}.zip'.format(
                self.start_date, self.finish_date, NOW
            )
        )
        logger.info(u'Пакуем файлы отчётов по терминалам в архив - {}'.format(
            output_file))
        compress_to_zip(input_files, output_file)

    def delete_tmp_gs_report_files(self, files):
        logger.info(u'Удаляем файлы отчётов по терминалам вне архива')
        for f in files:
            os.remove(f)

    def close(self):
        logger.info(u'Закрываем BOC')
        self.menu.MenuSelect(u'Файл->Выход')

    def run(self):
        self.menu.Wait('ready')
        self.login()
        self.fill_clients()
        self.fill_companies()
        self.fill_oracle()
        self.form_mainbase_and_svodbase()
        if self.need_archive_mainbase_and_svodbase:
            self.make_base_zip()
        if self.need_stat:
            self.form_mainlnr()
            self.calc_stat()
            self.make_stat_zip()
        if self.need_client_reports:
            self.make_client_reports()
            self.make_client_reports_zip()
        if self.need_gs_report:
            input_files = self.make_gs_report()
            self.make_gs_report_zip(input_files)
            self.delete_tmp_gs_report_files(input_files)
        self.close()
        logger.info(u'Все операции успешно завершены')


def main():
    today = datetime.date.today()
    this_month_first_day = datetime.date(today.year, today.month, 1)
    previous_month_first_day = (
        this_month_first_day - datetime.timedelta(days=1)).replace(day=1)

    parser = argparse.ArgumentParser()
    parser.add_argument('-v', '--version', action='version',
                        version='%(prog)s ' + __version__)
    parser.add_argument('boc_path', type=valid_file, help=u'Путь до BOC')
    parser.add_argument('-nbd', '--bd_name', help=u'Сервер БД')
    parser.add_argument('-ubd', '--bd_user', help=u'Пользователь БД')
    parser.add_argument('-pbd', '--bd_psw', help=u'Пароль БД')
    parser.add_argument(
        '-eo', help=u'Эксплуатирующая организация (по умолчанию: %(default)s)',
        default=u'Ульяновский филиал ООО "Татнефть-АЗС Центр"'
    )
    parser.add_argument(
        '-sd', '--start_date', type=valid_date, default=previous_month_first_day,
        help=u'Начальная дата в формате ГГГГ-ММ-ДД (по умолчанию: %(default)s)'
    )
    parser.add_argument(
        '-fd', '--finish_date', type=valid_date, default=this_month_first_day,
        help=u'Конечная дата в формате ГГГГ-ММ-ДД (по умолчанию: %(default)s)'
    )
    parser.add_argument(
        '-zms', '--zip_mainbase_svodbase', default=False, action='store_true',
        help=u'Запаковать в архив MAINBASE.dbf и SVODBASE.dbf'
    )
    parser.add_argument(
        '-zst', '--zip_stat', default=False, action='store_true',
        help=u'Запаковать в архив Stat456 и klients.dbf'
    )
    parser.add_argument(
        '-zcr', '--zip_client_reports', default=False, action='store_true',
        help=u'Запаковать в архив AllReports'
    )
    parser.add_argument(
        '-zgr', '--zip_gs_report', default=False, action='store_true',
        help=u'Запаковать в архив "Отчет для сверки с АЗС"'
    )

    args = parser.parse_args()
    logger.info(u'Переданные аргументы: {}'.format(args))

    working_dir = os.path.dirname(args.boc_path)
    os.chdir(working_dir)

    boc_auto = BocAutoGui(**vars(args))
    try:
        boc_auto.run()
    except:
        logger.exception(u'Неожиданное завершение программы')
        boc_auto.make_top_window_screenshot()
        boc_auto.kill_process()
        raise


if __name__ == '__main__':
    main()
