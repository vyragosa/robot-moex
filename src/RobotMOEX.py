import logging
from datetime import datetime

import pandas as pd

import src.MailActivity as mail_act
import src.OSActivity as OS_act
import src.constants as const
from src.BrowserActivity import BrowserActivity
from src.ExcelActivity import ExcelActivity


class RobotMOEX():
    def __init__(self):
        """
        Робот по получению данных с сайта moex.com.
        """
        self.curr_date = datetime.today().strftime('%d.%m.%Y')
        OS_act.create_folder(const.BASE_DIR)
        self.browser_act = BrowserActivity()
        self.browser_act.open_page(url=const.BASE_URL, max_attempts=3, wait_time=60)
        self.excel_act = ExcelActivity(filepath=f"{const.BASE_DIR}{self.curr_date}.xlsx")

    def shutdown(self):
        """
        Заверщение работы робота
        """
        self.excel_act.close()
        self.browser_act.close_browser()
        OS_act.remove_folder(const.BASE_DIR)

    def _get_dataframe(self, title: tuple[str, str, str]) -> pd.DataFrame:
        """
        Создает таблицу, полученную с сайта
        :param title: Заголовок таблицы
        :return: Таблица
        """
        return pd.DataFrame(data={
            title[0]: self.browser_act.get_texts(r'//tbody//tr/td[1]'),
            title[1]: [0 if rate == '-' else float(rate) for rate in
                       self.browser_act.get_texts(r'//tbody//tr/td[4]')],
            title[2]: self.browser_act.get_texts(r'//tbody//tr/td[5]'),
        })

    def _prepare_data(self) -> list:
        """
        Предподготовка данных. Осуществляет вход на сайт и сбор данных
        :return: Таблицы с данными курса
        """
        self.browser_act.click(r'/html/body/div[1]/div/div/div[3]/div/header/div[4]/div/div[2]/button')  # клик по меню
        self.browser_act.click('//body//header/div[5]/div[2]/div/div/div/ul/li[2]/a')  # клик по Срочный рынок

        if self.browser_act.element_exists(
                '//div[2]/div/div/div/div[@class="disclaimer"]/div[2]', ):  # Проверка на всплывающее окно с ПС
            self.browser_act.click(r'//div[2]/div/div/div/div/div[1]/div[@class="disclaimer__buttons"]/a', timeout=10,
                                   wait_for_ready=False)  # Клик по согласен
        self.browser_act.click(
            r'//div[@class="left-menu__dropdown "]//a[@href="/ru/derivatives/currency-rate.aspx?currency=USD_RUB"]')  # клик по индикативные курсы

        tableUSD = self._get_dataframe(title=("Дата USD/RUB", "Курс USD/RUB", "Время USD/RUB"))
        self.browser_act.click(r'//body//form/div[1]/div[1]')  # нажатие на селектор выбора курса
        self.browser_act.click(r'//div[5]/div[3]/div[1]/div[3]')  # нажатие на курс JPY
        tableJPY = self._get_dataframe(title=("Дата JPY/RUB", "Курс JPY/RUB", "Время JPY/RUB"))
        return tableUSD, tableJPY

    def _excel_parse(self, tableUSD: pd.DataFrame, tableJPY: pd.DataFrame) -> None:
        """
        Работа с файлом excel. Осуществляет запись данных в файл, изменение стиля на финансовый,
        добавление формул деления и автосуммы, настройку автощирины и сохранение
        :param tableUSD: Таблица для USD/RUB
        :param tableJPY: Таблица для JPY/RUB
        """
        self.excel_act.write_dataframe(tableUSD, start_col=1)
        self.excel_act.write_dataframe(tableJPY, start_col=4)
        self.excel_act.change_number_format(col="B")
        self.excel_act.change_number_format(col="E")
        self.excel_act.add_formulae(col="G")
        self.excel_act.autofit()
        self.excel_act.save()

    @staticmethod
    def _row_declension(row: int) -> str:
        """
        Получение правильного склонения слово Строка.
        :param row: номер строки
        :return: Строка в правильном склонении
        """
        if row % 10 == 1:
            return "строка"
        if row % 10 in [2, 3, 4]:
            return "строки"
        return "строк"

    def _main_activity(self, tableUSD: pd.DataFrame, tableJPY: pd.DataFrame) -> None:
        """
        Включает работу с данными с сайта и отправку письма по почте
        :param tableUSD: Таблица для USD/RUB
        :param tableJPY: Таблица для JPY/RUB
        """
        self._excel_parse(tableUSD, tableJPY)
        mail_act.send_mail(to=const.MAIL_TO,
                           subject=f"Индикативные курсы валют от {self.curr_date}",
                           body=f"Осуществлен сбор данных индикативного курса валют USD/RUB и JPY/RUB за {self.curr_date}."
                                f"В файле содержится {self.excel_act.get_rows()} {self._row_declension(self.excel_act.get_rows())}.",
                           attachment=f"{const.BASE_DIR}{self.curr_date}.xlsx")

    def run(self):
        """
        Основной контейнер Содержится главные блоки робота. Блок завершения работы.
        """
        try:
            tableUSD, tableJPY = self._prepare_data()
            self._main_activity(tableUSD, tableJPY)
        except Exception as e:
            logging.error(e)
        finally:
            self.shutdown()
