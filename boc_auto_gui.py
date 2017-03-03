# -*- coding: utf-8 -*-

__version__ = '0.3.2'

import datetime
import errno
import logging
import os
import zipfile
import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import COMMASPACE

import configargparse
import dbf
from pywinauto.application import Application, AppStartError
from pywinauto.timings import WaitUntil  # , Timings

# python 2 compatibility
try:
    FileNotFoundError
except NameError:
    FileNotFoundError = IOError

NOW = datetime.datetime.now()
#FILE_DIR_PATH = os.path.dirname(os.path.realpath(__file__))
FILE_DIR_PATH = os.path.dirname(os.path.abspath(__file__))


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
    format=u'[%(asctime)s] [%(levelname)-8s] %(message)s',
    level=logging.INFO)
################################################################################


def get_date_via_dialog(date_, start=True):
    today = datetime.date.today()
    choosen_date = False
    while not choosen_date:
        msg = (
            u'Введите {} дату в формате ГГГГ-ММ-ДД (по умолчанию: {}) '
            u'или только число для выбора даты в текущем месяце: '.format(
                u'начальную' if start else u'конечную', date_)
        )
        in_str = input(msg)
        if not in_str:
            choosen_date = date_
        else:
            try:
                choosen_date = datetime.datetime.strptime(
                    in_str, '%Y-%m-%d').date()
            except ValueError:
                try:
                    choosen_date = today.replace(day=int(in_str))
                except ValueError:
                    print(u'Неверный формат ввода')
    return choosen_date


def create_dir_if_not_exists(dir):
    try:
        os.makedirs(dir)
    except OSError as exception:
        if exception.errno != errno.EEXIST:
            raise


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
                error_msg = '{} - doesnot exists'.format(path)
                raise FileNotFoundError(error_msg)
    return output_file


def get_dbf_empty_field_value(table, fieldname):
    field_length = table.field_info(fieldname)[1]
    return ' ' * field_length


def valid_file(value):
    if os.path.isfile(value):
        return os.path.abspath(value)
    raise configargparse.ArgumentTypeError('{} is not a file'.format(value))


def valid_date(value):
    try:
        return datetime.datetime.strptime(value, '%Y-%m-%d').date()
    except ValueError:
        msg = 'Invalid date: "{}".'.format(value)
        raise configargparse.ArgumentTypeError(msg)


def create_mail(from_addr, to_addr, subject, message, files=None, reply_addr=None):
    msg = MIMEMultipart()
    msg['From'] = from_addr
    if isinstance(to_addr, (list, tuple)):
        to_addr = COMMASPACE.join(to_addr)
    msg['To'] = to_addr
    msg['Subject'] = subject
    if reply_addr:
        msg['reply-to'] = reply_addr
    msg.attach(MIMEText(message, 'plain'))
    if files:
        if not isinstance(files, (tuple, list)):
            files = [files]
        for f in files:
            with open(f, 'rb') as att:
                part = MIMEApplication(
                    att.read(),
                    Name=os.path.basename(f)
                )
                part['Content-Disposition'] = (
                    u'attachment; filename="{}"'.format(os.path.basename(f))
                )
                msg.attach(part)
    return msg


def send_email(host, port, from_addr, password, mails_with_destination):
    '''
    mails_with_destination - список формата
    [{'to_addr_list': to_addr_list, 'mail': mail}, ]
    '''
    # XXX timeout?
    server = smtplib.SMTP_SSL(host, port)
    try:
        server.login(from_addr, password)
        for msg in mails_with_destination:
            subject = msg['mail']['Subject']
            logger.info(u'Отправляем письмо - {}'.format(subject))
            try:
                server.sendmail(
                    from_addr, msg['to_addr_list'], msg['mail'].as_string())
            except:
                # XXX
                logger.exception(
                    u'При отправке письма {} произошла ошибка'.format(subject))
    finally:
        server.quit()


class BocAutoGui():
    def __init__(self, boc_path, bd_name, bd_user, bd_psw, eo,
                 mail_host, mail_port, from_addr, mail_password, bcc_addr,
                 reply_addr, to_addr_bases, to_addr_client_reports,
                 to_addr_gs_report, to_addr_stat,
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

        self.mail_host = mail_host
        self.mail_port = mail_port
        self.from_addr = from_addr
        self.mail_password = mail_password
        self.bcc_addr = bcc_addr
        self.reply_addr = reply_addr
        self.to_addr_bases = to_addr_bases
        self.to_addr_client_reports = to_addr_client_reports
        self.to_addr_gs_report = to_addr_gs_report
        self.to_addr_stat = to_addr_stat
        self.message = u'\n'.join([
            u'Файл вложен и отправлен автоматически.',
            u'Просьба отвечать на адрес {}'.format(self.reply_addr),
            u'---',
            u'Ульяновский филиал ООО "Процессинговый Центр"'
        ])

        self.mails = list()

        try:
            self.app = Application().Start(cmd_line=boc_path)
        except AppStartError:
            logger.exception(u'Не удалось запустить приложение')
            raise
        self.menu = self.app['TOraTypeMain']
        self.output_dir = os.path.join(FILE_DIR_PATH, 'output')
        create_dir_if_not_exists(self.output_dir)
        #Timings.Slow()  # Уведичиваем таймауты

    def kill_process(self):
        logger.warning(u'Убиваем процесс...')
        self.app.Kill_()

    def make_top_window_screenshot(self):
        img_path = os.path.join(
            self.output_dir,
            'error_{:%Y%m%d_%H%M%S}.png'.format(datetime.datetime.now())
        )
        try:
            self.app.top_window_().CaptureAsImage().save(img_path)
            logger.info(u'Скриншот с ошибкой - {}'.format(img_path))
        except AttributeError:
            logger.warning(
                u'Похоже не установлен PIL/Pillow. '
                u'Невозможно сделать скриншот'
            )
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
                # XXX косяк с кодировкой во 2 питоне
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
        self.menu.Wait('ready')
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
        clients_window[u'ОК'].Wait('ready')
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
                    self.start_date, self.finish_date, NOW, i+1)
            )
            output_path = os.path.join(self.output_dir, filename)
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
            self.output_dir,
            'base456_{:%Y%m%d}-{:%Y%m%d}_{:%Y%m%d%H%M%S}.zip'.format(
                self.start_date, self.finish_date, NOW)
        )
        logger.info(
            u'Пакуем файлы MAINBASE.dbf и SVODBASE.dbf в архив - {}'.format(
                output_file)
        )
        return compress_to_zip(input_files, output_file)

    def make_stat_zip(self):
        input_files = (
            os.path.join(self.boc_dir_path, 'Stat456'),
            os.path.join(self.boc_dir_path, 'klients.dbf'),
        )
        output_file = os.path.join(
            self.output_dir,
            'stat456_{:%Y%m%d}-{:%Y%m%d}_{:%Y%m%d%H%M%S}.zip'.format(
                self.start_date, self.finish_date, NOW)
        )
        logger.info(
            u'Пакуем файлы Stat456 и klients.dbf в архив - {}'.format(
                output_file)
        )
        return compress_to_zip(input_files, output_file)

    def make_client_reports_zip(self):
        input_files = os.path.join(self.boc_dir_path, 'AllReports')
        output_file = os.path.join(
            self.output_dir,
            'reports_{:%Y%m%d}-{:%Y%m%d}_{:%Y%m%d%H%M%S}.zip'.format(
                self.start_date, self.finish_date, NOW)
        )
        logger.info(u'Пакуем файлы AllReports в архив - {}'.format(
            output_file))
        return compress_to_zip(input_files, output_file)

    def make_gs_report_zip(self, input_files):
        output_file = os.path.join(
            self.output_dir,
            'gs_report_{:%Y%m%d}-{:%Y%m%d}_{:%Y%m%d%H%M%S}.zip'.format(
                self.start_date, self.finish_date, NOW)
        )
        logger.info(u'Пакуем файлы отчётов по терминалам в архив - {}'.format(
            output_file))
        return compress_to_zip(input_files, output_file)

    def delete_tmp_gs_report_files(self, files):
        logger.info(u'Удаляем файлы отчётов по терминалам вне архива')
        for f in files:
            os.remove(f)

    def add_mail_to_delivery(self, file_, to_addr):
        if isinstance(file_, (tuple, list)):
            subject = os.path.basename(file_[0]) + ' & other'
        else:
            subject = os.path.basename(file_)
        try:
            mail = create_mail(
                self.from_addr,
                to_addr,
                subject,
                self.message,
                file_,
                reply_addr=self.reply_addr
            )
        except:
            msg = u'Не удалось сформировать письмо с вложением - {}'.format(
                subject)
            logger.exception(msg)
            raise
        logger.info(u'Добавляем {} в рассылку'.format(subject))
        self.mails.append({
            'to_addr_list': to_addr+self.bcc_addr,
            'mail': mail
        })

    def send_mails(self):
        logger.info(u'Начинаем рассылку')
        send_email(
            self.mail_host,
            self.mail_port,
            self.from_addr,
            self.mail_password,
            self.mails
        )

    def close(self):
        logger.info(u'Закрываем BOC')
        self.menu.MenuSelect(u'Файл->Выход')

    def run(self):
        # XXX Желательно выполнять операции, несвязанные с GUI, после закрытия окна BOC.
        self.login()
        self.fill_clients()
        self.fill_companies()
        self.fill_oracle()
        self.form_mainbase_and_svodbase()
        if self.need_archive_mainbase_and_svodbase:
            file_ = self.make_base_zip()
            self.add_mail_to_delivery(file_, self.to_addr_bases)
        if self.need_stat:
            self.form_mainlnr()
            self.calc_stat()
            file_ = self.make_stat_zip()
            self.add_mail_to_delivery(file_, self.to_addr_stat)
        if self.need_client_reports:
            self.make_client_reports()
            file_ = self.make_client_reports_zip()
            self.add_mail_to_delivery(file_, self.to_addr_client_reports)
        if self.need_gs_report:
            input_files = self.make_gs_report()
            file_ = self.make_gs_report_zip(input_files)
            self.add_mail_to_delivery(file_, self.to_addr_gs_report)
            self.delete_tmp_gs_report_files(input_files)
        self.close()
        if self.mails:
            self.send_mails()
        logger.info(u'Все операции успешно завершены')


def main():
    today = datetime.date.today()
    this_month_first_day = datetime.date(today.year, today.month, 1)
    previous_month_first_day = (
        this_month_first_day - datetime.timedelta(days=1)).replace(day=1)
    defaults = {
        'ekspl_org': u'Ульяновский филиал ООО "Татнефть-АЗС Центр"',
        'start_date': previous_month_first_day,
        'finish_date': this_month_first_day,
        'zip_mainbase_svodbase': False,
        'zip_stat': False,
        'zip_client_reports': False,
        'zip_gs_report': False,
    }

    parser = configargparse.ArgumentParser()
    parser.set_defaults(**defaults)
    parser.add_argument('-v', '--version', action='version',
                        version='%(prog)s ' + __version__)
    parser.add_argument('-c', '--config', required=False, is_config_file=True,
                        help='Файл конфигурации')
    parser.add_argument('-b', '--boc_path', type=valid_file, required=True,
                        help=u'Путь до BOC')
    parser.add_argument('-nbd', '--bd_name', help=u'Сервер БД')
    parser.add_argument('-ubd', '--bd_user', help=u'Пользователь БД')
    parser.add_argument('-pbd', '--bd_psw', help=u'Пароль БД')
    parser.add_argument(
        '-eo', '--ekspl_org',
        help=u'Эксплуатирующая организация (по умолчанию: %(default)s)',
    )
    parser.add_argument(
        '-mh', '--mail_host',
        help=u'Хост электронной почты (по умолчанию: %(default)s)',
    )
    parser.add_argument(
        '-mp', '--mail_port', type=int,
        help=u'Порт для почты (по умолчанию: %(default)s)',
    )
    parser.add_argument(
        '-fa', '--from_addr', metavar='EMAIL',
        help=u'Почта с которой производим рассылку (по умолчанию: %(default)s)',
    )
    parser.add_argument(
        '-mpsw', '--mail_password',
        help=u'Пароль от почты (по умолчанию: %(default)s)',
    )
    parser.add_argument(
        '-ba', '--bcc_addr', nargs='*', metavar='EMAIL',
        help=u'Почта для отправки "слепой" копии (по умолчанию: %(default)s)',
    )
    parser.add_argument(
        '-ra', '--reply_addr', metavar='EMAIL',
        help=u'Почта для ответа (по умолчанию: %(default)s)',
    )
    parser.add_argument(
        '-ab', '--to_addr_bases', nargs='*', metavar='EMAIL',
        help=u'Почта для отправки базы (по умолчанию: %(default)s)',
    )
    parser.add_argument(
        '-ac', '--to_addr_client_reports', nargs='*', metavar='EMAIL',
        help=u'Почта для отправки отчета по клиентам (по умолчанию: %(default)s)',
    )
    parser.add_argument(
        '-ag', '--to_addr_gs_report', nargs='*', metavar='EMAIL',
        help=u'Почта для отправки отчета по АЗС (по умолчанию: %(default)s)',
    )
    parser.add_argument(
        '-as', '--to_addr_stat', nargs='*', metavar='EMAIL',
        help=u'Почта для отправки статистики (по умолчанию: %(default)s)',
    )
    parser.add_argument(
        '-sd', '--start_date', type=valid_date,
        help=u'Начальная дата в формате ГГГГ-ММ-ДД (по умолчанию: %(default)s)'
    )
    parser.add_argument(
        '-fd', '--finish_date', type=valid_date,
        help=u'Конечная дата в формате ГГГГ-ММ-ДД (по умолчанию: %(default)s)'
    )
    parser.add_argument(
        '-zms', '--zip_mainbase_svodbase', action='store_true',
        help=u'Запаковать в архив MAINBASE.dbf и SVODBASE.dbf'
    )
    parser.add_argument(
        '-zst', '--zip_stat', action='store_true',
        help=u'Запаковать в архив Stat456 и klients.dbf'
    )
    parser.add_argument(
        '-zcr', '--zip_client_reports', action='store_true',
        help=u'Запаковать в архив AllReports'
    )
    parser.add_argument(
        '-zgr', '--zip_gs_report', action='store_true',
        help=u'Запаковать в архив "Отчет для сверки с АЗС"'
    )
    parser.add_argument(
        '-npd', '--no_period_dialog', action='store_true',
        help=u'Не требовать подтверждение временного периода'
    )

    args = parser.parse_args()

    if not (args.no_period_dialog and all([args.start_date, args.finish_date])):
        args.finish_date = get_date_via_dialog(args.finish_date, False)
        args.start_date = get_date_via_dialog(args.finish_date.replace(day=1))

    logger.info(u'Переданные аргументы: {}'.format(args))

    working_dir = os.path.dirname(args.boc_path)
    os.chdir(working_dir)

    boc_auto = BocAutoGui(
        boc_path=args.boc_path,
        bd_name=args.bd_name,
        bd_user=args.bd_user,
        bd_psw=args.bd_psw,

        mail_host=args.mail_host,
        mail_port=args.mail_port,
        from_addr=args.from_addr,
        mail_password=args.mail_password,
        bcc_addr=args.bcc_addr,
        reply_addr=args.reply_addr,
        to_addr_bases=args.to_addr_bases,
        to_addr_client_reports=args.to_addr_client_reports,
        to_addr_gs_report=args.to_addr_gs_report,
        to_addr_stat=args.to_addr_stat,

        eo=args.ekspl_org,
        start_date=args.start_date,
        finish_date=args.finish_date,
        zip_mainbase_svodbase=args.zip_mainbase_svodbase,
        zip_stat=args.zip_stat,
        zip_client_reports=args.zip_client_reports,
        zip_gs_report=args.zip_gs_report
    )
    try:
        boc_auto.run()
    except:
        logger.exception(u'Неожиданное завершение программы')
        boc_auto.make_top_window_screenshot()
        boc_auto.kill_process()
        raise


if __name__ == '__main__':
    main()

# XXX если в eksplorg не задана Эксплаутирующая организация для терминала будет вызвана ошибка
# Предлагать заполнять это поле вручную?
# Уточнить обязательные/необязательные аргументы для скрипта
