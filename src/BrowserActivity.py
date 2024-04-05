import time
from selenium import webdriver
from selenium.common.exceptions import WebDriverException, TimeoutException, NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


class BrowserActivity:
    def __init__(self):
        self.chrome_options = webdriver.ChromeOptions()
        self.chrome_options.add_argument("--start-maximized")
        self.driver = webdriver.Chrome(options=self.chrome_options)

    def open_page(self, url: str, max_attempts: int = 0, wait_time: int = 0) -> None:
        """
        Открывает страницу в браузере Google Chrome. Если не удалось зайти, осуществляет повторный вход
        :param url: Адрес страницы
        :param max_attempts: Максимальное количество попыток входа на сайт
        :param wait_time: Время ожидания между перезаходами
        """
        for attempt in range(max_attempts):
            try:
                self.driver.get(url)
                break
            except WebDriverException as e:
                print(e)
                if attempt < max_attempts - 1:
                    print("Retrying...")
                    time.sleep(wait_time)
                else:
                    raise Exception(f"Failed to open page: {url}.")

    def click(self, element: str, wait_for_ready: bool = True, timeout: int = 60) -> None:
        """
        Нажимает на элемент в окне браузера
        :param element: XPath элемента
        :param wait_for_ready: Ожидание появления элементе
        :param timeout: Время ожидания появления элемента
        """
        try:
            if wait_for_ready:
                self.element_exists(element, wait_for_ready, timeout)
            self.driver.find_element(By.XPATH, element).click()
        except TimeoutException:
            raise Exception(f"Timed out waiting for element '{element}'.")
        except NoSuchElementException:
            raise Exception(f"Element '{element}' not found.")

    def element_exists(self, element: str, wait_for_ready: bool = True, timeout: int = 60) -> bool:
        """
        Проверка наличия элемента
        :param element: XPath элемента
        :param wait_for_ready: Ожидание появления элемента
        :param timeout: Время ожидания появления элемента
        :return: Наличие элемента на странице
        """
        try:
            if wait_for_ready:
                WebDriverWait(self.driver, timeout).until(EC.presence_of_element_located((By.XPATH, element)))
            return True
        except TimeoutException:
            return False

    def get_texts(self, element: str, wait_for_ready: bool = True, timeout: int = 60) -> list:
        """
        Возврат текста элемента
        :param element: XPath элемента
        :param wait_for_ready: Ожидание появления элемента
        :param timeout: Время ожидания появления элемента
        :return: Список значений текста элемента
        """
        try:
            if wait_for_ready:
                self.element_exists(element, wait_for_ready, timeout)
            contents = self.driver.find_elements(By.XPATH, element)
            return [content.text for content in contents]
        except TimeoutException:
            raise Exception(f"Timed out waiting for element '{element}' to be present.")
        except NoSuchElementException:
            raise Exception(f"Element '{element}' not found.")

    def close_browser(self):
        self.driver.quit()
