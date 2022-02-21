import glob
import os
import pandas as pd
import platform
import schedule
import time

from datetime import datetime
from dotenv import load_dotenv

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.utils import ChromeType

load_dotenv()


def download_reports(driver, duration_type, temp_dir, report_dir, year):
    """
    Скачивание отчетов и их последующая обработка

    :param driver: драйвер
    :param duration_type: вид плана
    :param temp_dir: временная папка для сохранения файлов
    :param report_dir: папка для сохранения отчетов
    :param year: финансовый год
    :return: список номеров закупок
    """

    wait = WebDriverWait(driver, 60)

    temp_dir_before = glob.glob(temp_dir + '/*.xls')   # содержимое временной папки для того, чтобы найти скаченный файл

    # Выбор необходимого вида плана и последующий клик на кнопку формирвоания отчета
    driver.find_element(
        By.XPATH,
        f'//sk-linear[@name="durationType"]//select/option[text()=" {duration_type} "]'
    ).click()
    driver.find_element(
        By.XPATH,
        '//button/span[@jhitranslate="planExecutionReport.download"]'
    ).click()

    try:
        # Ожидание окончания загрузки отчета
        wait.until(
            EC.element_to_be_clickable((
                By.XPATH,
                '//button/span[@jhitranslate="planExecutionReport.download"]'
            ))
        )

        # Проверка на наличие новых excel-файлов во временной папке
        while glob.glob(temp_dir + '/*.xls') == temp_dir_before:
            time.sleep(1)

        # Извлечение скаченного отчета
        temp_dir_after = glob.glob(temp_dir + '/*.xls')
        downloaded_file = list(set(temp_dir_after) - set(temp_dir_before))[-1]

        # Перемещение и переименование отчета
        report = f'{report_dir}/Отчет по исполнению плана от {datetime.now().strftime("%d-%m-%Y %H:%M:%S")}.xls'
        os.rename(downloaded_file, report)

    except Exception as e:
        with open('log.info', 'a+') as f:
            content = f' - Не удается выгрузить отчет по исполнению плана {year} - Перечень - {duration_type}\n'
            f.write(datetime.now().strftime("%d-%m-%Y %H:%M:%S") + content)
            return

    # Обработка excel-файла
    df = pd.read_excel(report, skiprows=1, engine='openpyxl')
    mapping = pd.read_excel('./mapping.xlsx', engine='openpyxl')

    # Сортировка
    df.sort_values(
        ['Номер строки плана закупок'],
        key=lambda x: x.str.extract('(\d+)').squeeze(),
        inplace=True
    )

    # Добавление нового столбца
    df.insert(df.columns.get_loc('Статус договора') + 1, 'Статус для свода', '')

    purchases = []
    for index, row in df.iterrows():
        purchase_status = row['Статус закупки']
        contract_status = row['Статус договора']

        # Поиск и запись неободимого статуса согласно маппингу
        statuses = mapping.loc[
            (mapping['Статус закупки'].isin([purchase_status])) & (mapping['Статус договора'].isin([contract_status])),
            'Для отчета КЦ КМГ'
        ].values

        df['Статус для свода'].values[index] = ' / '.join(statuses) or '-'

        if row['Номер закупки'] != '-':
            purchases.append(row['Номер закупки'])

    df.to_excel(report, index=False)

    return purchases


def download_purchase_documents(driver, purchase, temp_dir, document_dir):
    """
    Скачивание документов закупок

    :param driver: драйвер
    :param purchase: номер закупки
    :param temp_dir: временная папка для сохранения файлов
    :param document_dir: папка для сохранения документов
    :return: None
    """

    wait = WebDriverWait(driver, 10)

    # Переход на страницу списка закупок
    wait.until(EC.element_to_be_clickable((By.XPATH, '//a/span[@jhitranslate="layouts.advert"]'))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, '//a/span[@jhitranslate="layouts.advertList"]'))).click()

    # Поиск закупки по номеру
    purchase_input = driver.find_element(By.XPATH, '//sk-textbox[@name="advertNumber"]//input')
    purchase_input.clear()
    purchase_input.send_keys(purchase)
    driver.find_element(By.XPATH, '//button/span[@jhitranslate="eProcGatewayApp.advert.searchParam.search"]').click()

    # Переход на страницу просмотра закупки
    actions = wait.until(
        EC.element_to_be_clickable((
            By.XPATH,
            f'//tr[@data-index="{purchase}"]//button[@jhitranslate="eProcGatewayApp.advert.actions"]'
        ))
    )
    driver.execute_script("arguments[0].click();", actions)
    wait.until(
        EC.element_to_be_clickable((
            By.XPATH,
            f'//tr[@data-index="{purchase}"]//a[@jhitranslate="eProcGatewayApp.advert.goAdvert"]'
        ))
    ).click()

    temp_dir_before = glob.glob(temp_dir + '/*')   # содержимое временной папки для того, чтобы найти скаченный файл

    # Проверка наличия документов на странице закупки
    try:
        documents = wait.until(EC.presence_of_all_elements_located((By.XPATH, '//sk-file-link//a')))
    except TimeoutException:
        return

    # Обработка каждого документа на странице закупки
    for document in documents:
        document.click()

        try:
            # Проверка на наличие новых файлов во временной папке
            while glob.glob(temp_dir + '/*') == temp_dir_before:
                time.sleep(1)

            # Извлечение скаченного отчета
            temp_dir_after = glob.glob(temp_dir + '/*')
            downloaded_file = list(set(temp_dir_after) - set(temp_dir_before))[-1]

            # Перемещение и переименование отчета
            filename, extension = os.path.splitext(downloaded_file)
            os.rename(downloaded_file, f'{document_dir}/{purchase} - {document.text}{extension}')

        except Exception:
            with open('log.info', 'a+') as f:
                content = f' - Не удается выгрузить документ {purchase} - {document.text}\n'
                f.write(datetime.now().strftime("%d-%m-%Y %H:%M:%S") + content)
                return


def main():
    """
    Основной функционал

    :return: None
    """

    # Получение пути к рабочему столу в зависимости от ОС пользователя
    if platform.system() == 'Windows':
        desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
    else:
        desktop = os.path.join(os.path.join(os.path.expanduser('~')), 'Desktop')

    # Поиск / создание временной и рабочих папок
    temp_dir = os.path.join(desktop, 'testnewtender_temp')
    annual_dir = os.path.join(desktop, 'ГПЗ')
    long_term_dir = os.path.join(desktop, 'ДПЗ')
    document_dir = os.path.join(desktop, 'ИСЭЗ')

    for directory in [temp_dir, annual_dir, long_term_dir, document_dir]:
        if not os.path.exists(directory):
            os.makedirs(directory)

    # Инициализация драйвера
    options = webdriver.ChromeOptions()
    options.add_experimental_option('prefs', {
        'download.default_directory': temp_dir,
        'download.prompt_for_download': False,
        'download.directory_upgrade': True,
        'safebrowsing_for_trusted_sources_enabled': False,
        'safebrowsing.enabled': False
    })
    service = Service(ChromeDriverManager(chrome_type=ChromeType.CHROMIUM).install())

    driver = webdriver.Chrome(service=service, options=options)
    wait = WebDriverWait(driver, 10)

    # Запуск драйера
    try:
        driver.get(os.getenv('HOST'))

        wait.until(EC.element_to_be_clickable((By.XPATH, '//a[@jhitranslate="layouts.register"]'))).click()

        # Закрытие алерта браузера
        wait.until(EC.alert_is_present())
        alert = driver.switch_to.alert
        alert.dismiss()

        # Авторизация
        driver.find_element(By.XPATH, '//a/span[@jhitranslate="global.menu.account.loginWithoutEds"]').click()
        driver.find_element(By.XPATH, '//sk-textbox[@name="username"]//input').send_keys(os.getenv('LOGIN'))
        driver.find_element(By.XPATH, '//sk-passwordbox[@name="password"]//input').send_keys(os.getenv('PASSWORD'))
        driver.find_element(By.XPATH, '//button[@jhitranslate="login.form.button"]').click()

        # Переход на страницу отчетов
        wait.until(
            EC.element_to_be_clickable((
                By.XPATH,
                '//a/span[@jhitranslate="reports.reports"]'
            ))
        ).click()
        wait.until(
            EC.element_to_be_clickable((
                By.XPATH,
                '//a/span[@jhitranslate="layouts.planExecutionReport"]'
            ))
        ).click()

        # Ввод данных для фильтрации отчетов
        previous_year = datetime.now().year - 1   # выбор предыдушего года
        year = driver.find_element(By.XPATH, '//sk-numberbox[@name="planYear"]//input')
        year.clear()
        year.send_keys(previous_year)
        driver.find_element(By.XPATH, '//sk-linear[@name="planType"]//select/option[text()=" Основной план "]').click()

        # Скачивание и обраотка отчетов
        annual_report = download_reports(driver, 'Годовой', temp_dir, annual_dir, previous_year)
        long_term_report = download_reports(driver, 'Долгосрочный', temp_dir, long_term_dir, previous_year)

        # Скачивание документов из закупок
        if annual_report:
            for purchase in annual_report:
                download_purchase_documents(driver, purchase, temp_dir, document_dir)

        if long_term_report:
            for purchase in long_term_report:
                download_purchase_documents(driver, purchase, temp_dir, document_dir)

    except Exception as e:
        with open('log.info', 'a+') as f:
            content = f' - Произошла непредвиденная ошибка: {e}\n'
            f.write(datetime.now().strftime("%d-%m-%Y %H:%M:%S") + content)
            return

    finally:
        driver.close()

    os.rmdir(temp_dir)


if __name__ == '__main__':
    main()

    # Планирование запуска
    # schedule.every().day.at('13:30').do(main)
    # schedule.every().day.at('18:30').do(main)
    #
    # while True:
    #     schedule.run_pending()
    #     time.sleep(1)
