import datetime

import openpyxl as ox
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting.rule import Rule, ColorScaleRule, CellIsRule, FormulaRule
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.styles.borders import Side
from openpyxl.styles import PatternFill, Font, Border, Alignment
import time
from sklearn import tree
import pandas as pd
from bs4 import BeautifulSoup
import requests
import numpy as np
import re
import datetime as dt
from tqdm.auto import tqdm, trange
from copy import copy
from natasha import AddrExtractor, MorphVocab
import traceback
import pymorphy2
import pymorphy2_dicts_ru
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
import logging
import os
# для работы progress_apply
tqdm.pandas(desc="In Progress")


def is_lot_finished(driver, lot_url):
    lot_status = 0
    final_price = np.nan
    try:
        driver.get(lot_url)
        try:
            wait = WebDriverWait(driver, 10)
            wait.until(EC.presence_of_element_located((By.XPATH, #visibility_of_element_located
             "//div[@class='notice-pub-title-and-subtitle']")))
#            driver.implicitly_wait(1)
            if driver.find_element(By.XPATH, "//div[@class='notice-status with-background amsterdam']"):
                if driver.find_element(By.XPATH, "//div[@class='notice-status with-background amsterdam']").text.strip() == 'Завершено':
                    try:
                        wait = WebDriverWait(driver, 1)
                        wait.until(EC.presence_of_element_located((By.XPATH,
                         "//div[@class='lot-collapsed_header']")))
#                        driver.implicitly_wait(1)
                        try:
                            not_succeed_el = driver.find_elements(By.XPATH, "//div[@class='protocol-lot-status error']")
                        except:
                            pass
                        if not_succeed_el:
                            if not_succeed_el.text.strip() == 'Не состоялся':
                                lot_status = -1
                    #---> перенос второго if

                    # except (Exception,):
                    #     pass
                    #
                    # try:
                    #     wait = WebDriverWait(driver, 10)
                    #     waintil(EC.visibility_of_element_located((By.XPATH,
                    #      "//div[@clt.uass='lot-collapsed_header']")))
 #                       driver.implicitly_wait(1)
                        elif driver.find_element(By.XPATH, "//div[@class='protocol-lot-status amsterdam']"):
                            plsa_element = driver.find_element(By.XPATH, "//div[@class='protocol-lot-status amsterdam']")
                            if plsa_element.text.strip() == 'Состоялся':
                                lot_status = 1
                                #wait = WebDriverWait(driver, 10)
                                try:
                                    wait.until(EC.presence_of_element_located((By.XPATH,
                                    '//div[@class="lot-protocols"]/div/app-inline-icon/app-icon-chevron-right/*[name()="svg"]'))).click()
                                except:
                                    pass
                                wait.until(EC.presence_of_element_located((By.XPATH,
                                 '//span[contains(@class,"button__label") and contains(text(), "итог") '
                                 'and contains(text(), "аукцион")]'))).click()
                                # wait.until(EC.visibility_of_element_located((By.XPATH,
                                #  '//button[contains(.,"Информация")]'))).click()
                                final_price = int(re.sub(' ', '', wait.until(EC.presence_of_element_located((By.XPATH,
                                 '//div[@class="content__result__bet__value"]'))).text.split(',')[0]))
                    except:
                        pass

        except:
            pass

    except (Exception,):
        pass

    return (driver, lot_status, final_price)


def go_to_procedure_detail(driver, tab_name):
    element = driver.find_element(By.XPATH, "//div[text()=tab_name]")
    body = driver.find_element_by_css_selector('body')

    try:
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//div[text()=@tab_name]")))
        # body.send_keys(Keys.ARROW_DOWN)
        # body.send_keys(Keys.ARROW_DOWN)
        # body.send_keys(Keys.ARROW_DOWN)
        # body.send_keys(Keys.ARROW_DOWN)
        # body.send_keys(Keys.ARROW_DOWN)
        #        time.sleep(1)
        element.click()
    except (Exception,):
        try:
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//div[text()=@tab_name]")))
            body.send_keys(Keys.ARROW_DOWN)
            body.send_keys(Keys.ARROW_DOWN)
            body.send_keys(Keys.ARROW_DOWN)
            #            time.sleep(1)
            element.click()
        except:
            try:
                body.send_keys(Keys.ARROW_UP)
                body.send_keys(Keys.ARROW_UP)
                body.send_keys(Keys.ARROW_UP)
                body.send_keys(Keys.ARROW_UP)
                body.send_keys(Keys.ARROW_UP)
                body.send_keys(Keys.ARROW_UP)
                #                time.sleep(1)
                element.click()
            except:
                try:
                    WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//div[text()=@tab_name]")))
                    #                    time.sleep(1)
                    element.click()
                except:
                    print('Dont click')
                    logging.info('Dont click')
    return driver

def parse_lot(driver, url):
    """ Функция получает на вход URL лота с сайта investmoscow.ru, парсит данный лот и возвращает pd.Series
        с характеристиками лота.
        :param url: url лота
        :return: возвращает Pandas Series с данными по лоту.
        """

    # создаем словарь объекта
    obj_dict = {LOT_FIELDS[i]: np.nan for i in range(len(LOT_FIELDS))}
    driver.get(url)
    wait = WebDriverWait(driver, 10)
    try:
        wait.until(
            EC.visibility_of_element_located((By.XPATH,
                                          '//div[@class="tender__content"]')))
    except:
        driver.refresh()
        wait.until(
            EC.visibility_of_element_located((By.XPATH,
                                              '//div[@class="tender__content"]')))
    full_info = BeautifulSoup(driver.find_element(By.XPATH, '//div[@class="tender__content"]').get_attribute('innerHTML'),
                         "html.parser")

    # выбираем данные по площади и местоположению
    try:
        obj_dict['obj_square'] = float(
            str.replace(full_info.find('span', class_="uid-text-accent uid-mr-16").text, ',', '.').split(' ')[0])
    except Exception as exc:
        print('Ошибка определения площади объекта ', url, exc)
        logging.info(f"Лот {url} - Ошибка определения площади объекта")

    try:
        obj_dict['lot_tag'] = full_info.find('div', class_="tender-card__header-label").text.strip()
    except (Exception,):
        obj_dict['lot_tag'] = np.nan
        print('Ошибка определения номера лота объекта ', url, exc)
        logging.info(f"Лот {url} - Ошибка определения номера лота объекта")

    try:
        obj_dict['addr_string'] = full_info.find('div', class_="tender-card__description").div.next_sibling.text
    except Exception as exc:
        print('Ошибка определения адреса объекта ', url, exc)
        logging.info(f"Лот {url} - Ошибка определения адреса объекта")


    obj_desc = full_info.find('div', class_="subject-table")
    # данные из блока "Сведения об объекте"
    # забираем только интересующие данные
    t1 = obj_desc.find_all('div', class_="subject-table-row")
    for block in t1:
        label_text = block.find('div', class_="subject-table-label").text.strip()
        if label_text == 'Этаж:':
            try:
                obj_dict['addr_floor'] = int(block.find('div', class_="subject-table-text").text)
            except Exception as exc:
                print(f"Лот {url} - Ошибка определения этажа объекта")
                logging.info(f"Лот {url} - Ошибка определения этажа объекта")

        elif label_text == 'Этажность дома:':
            try:
                obj_dict['total_floors'] = int(block.find('div', class_="subject-table-text").text)
            except Exception as exc:
                print(f"Лот {url} - Ошибка определения этажности объекта")
                logging.info(f"Лот {url} - Ошибка определения этажности объекта")

        elif label_text == 'Тип объекта:':
            obj_dict['object_type'] = block.find('div', class_="subject-table-text").text

        elif label_text == 'Кадастровый номер:':
            obj_dict['cadastr_num'] = block.find('div', class_="subject-table-text").text

        elif label_text == 'Номер квартиры:':
            try:
                obj_dict['flat_num'] = int(block.find('div', class_="subject-table-text").text)
            except Exception as exc:
                print(f"Лот {url} - Ошибка определения номера квартиры объекта")
                logging.info(f"Лот {url} - Ошибка определения номера квартиры объекта")

        elif label_text == 'Количество комнат:':
            try:
                obj_dict['qty_rooms'] = int(block.find('div', class_="subject-table-text").text)
            except Exception as exc:
                print(f"Лот {url} - Ошибка определения количества комнат объекта")
                logging.info(f"Лот {url} - Ошибка определения количества комнат объекта")
    # переключаемся на "Сведения о процедуре"

    driver = go_to_procedure_detail(driver, 'Сведения о процедуре')


    wait.until(
        EC.visibility_of_element_located((By.XPATH,
                                          '//div[@class="tender__content"]')))
    full_info = BeautifulSoup(
        driver.find_element(By.XPATH, '//div[@class="tender__content"]').get_attribute('innerHTML'),
        "html.parser")
    obj_desc = full_info.find('div', class_="subject-table")
    t1 = obj_desc.find_all('div', class_="subject-table-row")

    for block in t1:
        label_text = block.find('div', class_="subject-table-label").text.strip()

        if label_text == 'Начальная цена за объект:':
            try:
                obj_dict['start_price'] = int(re.sub(r"\D", "",
                                            block.find('div', class_="subject-table-text").text.split(
                                                         ',')[0]))
                obj_dict['start_price_m2'] = round(obj_dict['start_price'] / obj_dict['obj_square'], 2)
            except Exception as exc:
                print(f"Лот {url} - Ошибка определения стартовой цены объекта")
                logging.info(f"Лот {url} - Ошибка определения стартовой цены объекта")

        elif label_text == 'Размер задатка:':
            try:
                obj_dict['deposit'] = int(re.sub(r"\D", "",
                                            block.find('div', class_="subject-table-text").text.split(
                                                         ',')[0]))
            except Exception as exc:
                print(f"Лот {url} - Ошибка определения задатка объекта")
                logging.info(f"Лот {url} - Ошибка определения задатка объекта")

        elif label_text == 'Шаг аукциона:':
            try:
                obj_dict['auct_step'] = int(re.sub(r"\D", "",
                                            block.find('div', class_="subject-table-text").text.split(
                                                         ',')[0]))
            except Exception as exc:
                print(f"Лот {url} - Ошибка определения шага аукциона объекта")
                logging.info(f"Лот {url} - Ошибка определения шага аукциона объекта")

        elif label_text == 'Форма проведения:':
            obj_dict['auct_form'] = block.find('div', class_="subject-table-text").text

        elif label_text == 'Проведение торгов:':
            obj_dict['auct_date'] = pd.to_datetime(
                block.find('div', class_="subject-table-text").text.split()[0],
                infer_datetime_format=True, format='%d.%m.%Y', errors='coerce')

        elif label_text == 'Дата начала приёма заявок:':
            obj_dict['start_applications_date'] = pd.to_datetime(
                block.find('div', class_="subject-table-text").text.split()[0],
                infer_datetime_format=True, format='%d.%m.%Y', errors='coerce')

        elif label_text == 'Дата окончания приёма заявок:':
            obj_dict['finish_applications_date'] = pd.to_datetime(
                block.find('div', class_="subject-table-text").text.split()[0],
                infer_datetime_format=True, format='%d.%m.%Y', errors='coerce')

        elif label_text == 'Отбор участников:':
            obj_dict['participant_selection_date'] = pd.to_datetime(
                block.find('div', class_="subject-table-text").text.split()[0],
                infer_datetime_format=True, format='%d.%m.%Y', errors='coerce')

        elif label_text == 'Проведение торгов:':
            obj_dict['bidding_date'] = pd.to_datetime(
                block.find('div', class_="subject-table-text").text.split()[0],
                infer_datetime_format=True, format='%d.%m.%Y', errors='coerce')

        elif label_text == 'Подведение итогов:':
            obj_dict['results_date'] = pd.to_datetime(
                block.find('div', class_="subject-table-text").text.split()[0],
                infer_datetime_format=True, format='%d.%m.%Y', errors='coerce')

        elif label_text == 'Ссылка на ЭТП:':
            obj_dict['roseltorg_url'] = block.find('a')['href']

        elif label_text == 'Ссылка на torgi.gov.ru:':
            obj_dict['torgi_url'] = block.find('a')['href']


    # переключаемся на "Дополнительная информация"

    driver = go_to_procedure_detail(driver, 'Дополнительная информация')


    wait.until(
        EC.visibility_of_element_located((By.XPATH,
                                          '//div[@class="tender__content"]')))
    full_info = BeautifulSoup(
        driver.find_element(By.XPATH, '//div[@class="tender__content"]').get_attribute('innerHTML'),
        "html.parser")
    obj_desc = full_info.find('div', class_="subject-table")
    t1 = obj_desc.find_all('div', class_="subject-table-row")

    for block in t1:
        label_text = block.find('div', class_="subject-table-label").text.strip()

        if label_text == 'Начальная цена за объект:':
            try:
                obj_dict['start_price'] = int(re.sub(r"\D", "",
                                            block.find('div', class_="subject-table-text").text.split(
                                                         ',')[0]))
                obj_dict['start_price_m2'] = round(obj_dict['start_price'] / obj_dict['obj_square'], 2)
            except Exception as exc:
                print(f"Лот {url} - Ошибка определения стартовой цены объекта")
                logging.info(f"Лот {url} - Ошибка определения стартовой цены объекта")



    return driver, pd.Series(obj_dict)


def control_lot(driver, url):
    """ Функция получает на вход URL лота с сайта investmoscow.ru.
        :param url: url лота
        :return: возвращает Pandas Series с данными по лоту.
        """

    # создаем словарь объекта
    obj_dict = dict({'status': np.nan,
                     'final_price': np.nan,
                     'final_price_m2': np.nan,
                     'delta_price': np.nan
                     })

    driver.get(url)
    wait = WebDriverWait(driver, 10)
    try:
        wait.until(
            EC.visibility_of_element_located((By.XPATH,
                                          '//div[@class="tender__content"]')))
    except:
        driver.refresh()
        try:
            wait.until(
                EC.visibility_of_element_located((By.XPATH,
                                                  '//div[@class="tender__content"]')))
        except:
            pass
    try:
        obj_dict['status'] = driver.find_element(By.XPATH, '//div[@class="tender-status__text"]').text.strip()
    except:
        obj_dict['status'] = 'failed_to_parse'

    if obj_dict['status'] == 'ПРИЗНАНЫ СОСТОЯВШИМИСЯ':
        # переключаемся на "Сведения о процедуре"
        #---> версия до 17.12.22
        # wait.until(
        #     EC.element_to_be_clickable((By.XPATH,
        #                                 "//div[@class='subject-tabs']/ul/li[2]")))
        # element = driver.find_element(By.XPATH, "//div[@class='subject-tabs']/ul/li[2]")
        # try:
        #     element.click()
        # except:
        #     pass
        # ActionChains(driver).move_to_element(
        #     WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH,
        #                                       "//div[@class='subject-tabs']/ul/li[2]")))).click().perform()
        #-------<версия до 17.12.22

        # ---> версия переключения на сведения о процедуре от 17.12.22
        driver = go_to_procedure_detail(driver)

        wait.until(
            EC.visibility_of_element_located((By.XPATH,
                                              '//div[@class="tender__content"]')))
        full_info = BeautifulSoup(
            driver.find_element(By.XPATH, '//div[@class="tender__content"]').get_attribute('innerHTML'),
            "html.parser")
        obj_desc = full_info.find('div', class_="subject-table")
        t1 = obj_desc.find_all('div', class_="subject-table-row")
        for block in t1:
            label_text = block.find('div', class_="subject-table-label").text.strip()
            # при флаге прошедших - добавляется итоговая цена + цена за м2 + дельта

            if label_text == 'Итоговая цена:':
                obj_dict['final_price'] = int(
                    re.sub(r"\D", "", block.find('div', class_="subject-table-text").text.split(',')[0]))


        # try:
        #     obj_dict['final_price_m2'] = round(obj_dict['final_price'] / obj_dict['obj_square'], 2)
        #     obj_dict['delta_price'] = round(
        #         (obj_dict['final_price'] - obj_dict['start_price']) / obj_dict['start_price'], 4)
        # except (Exception,):
        #     obj_dict['final_price_m2'] = np.nan
        # if (obj_dict['final_price_m2'] == 0) or (pd.isnull(obj_dict['final_price_m2'])):
        #     obj_dict['final_price_m2'] = 'не состоялись(проверка)'
    return driver, pd.Series(obj_dict)


def primary_addr_normalize(line):
    """
    Функция первичной нормализации адреса. Получает на вход строку адреса и пытается ее распарсить,
    не углубляясь в детали. Срабатывает на большую часть стандартно оформленных адресов.
    Возвращает список из строки с указанием города/улицы и строки с указанием дома/корпуса/строения
    """
    adr_street = ''
    adr_house = ''
    adr_floor = np.nan
    for _, paramValue in enumerate(addr_extractor.find(line).fact.parts):
        p0 = addr_extractor.find(line).fact.parts
    try:
        for p in p0:
            if p.type == 'город':
                if p.value != 'Москва' and p.value != 'Москвы':
                    adr_street += str(', г. ' + p.value)
            if p.type == 'деревня':
                adr_street += str(', д. ' + p.value)
            if p.type == 'микрорайон':
                adr_street += str(', мкр. ' + p.value)
            if p.type == 'бульвар':
                adr_street += str(', б-р. ' + p.value)
            if p.type == 'дачный поселок':
                adr_street += str(', дп. ' + p.value)
            if p.type == 'шоссе':
                adr_street += str(', ш. ' + p.value)
            if p.type == 'улица':
                adr_street += str(', ул. ' + p.value)
            if p.type == 'переулок':
                adr_street += str(', пер. ' + p.value)
            if p.type == 'квартал':
                adr_street += str(', кв-л. ' + p.value)
            if p.type == 'проезд':
                adr_street += str(', проезд. ' + p.value)
            if p.type == 'проспект':
                adr_street += str(', пр-кт. ' + p.value)
            if p.type == 'аллея':
                adr_street += str(', аллея. ' + p.value)
            if p.type == 'площадь':
                adr_street += str(', пл. ' + p.value)
            if p.type == 'набережная':
                adr_street += str(', наб. ' + p.value)
            if p.type == 'село':
                adr_street += str(', с. ' + p.value)
            if p.type == 'дом':
                adr_house = str(', д. ' + p.value)
            if p.type == 'корпус':
                adr_house = adr_house + str(', к. ' + p.value)
            if p.type == 'строение':
                adr_house = adr_house + str(', стр. ' + p.value)
            if p.type == 'этаж':
                adr_floor = p.value
        adr_street = 'г. Москва' + adr_street
        ret = [adr_street, adr_house, adr_floor]
    except:
        ret = ['не распознан', 'не распознан', 'не распознан']
    return ret


def secondary_addr_normalize(line):
    """
    Функция углубленной нормализации адреса. Получает на вход строку адреса и пытается ее распарсить,
    при этом проверяет вхождение слов из словаря conv_dictionary.xlsx и заменяет на соответствующие,
    переносит нумерацию улиц в конец строки (например, 1-я Трудовая должна быть "Трудовая 1-я ул."),
    обрабатывает названия Большой, Нижний и т.д., заменяет букву "ё" на "е".
    Так как делает больше проверок, запускается на части датафрейма, которую не удалось распарсить певично.
    Возвращает список из строки с указанием города/улицы и строки с указанием дома/корпуса/строения
    """
    # morph_vocab = MorphVocab()
    # addr_extractor = AddrExtractor(morph_vocab)
    adr_street = ''
    adr_house = ''
    adr_floor = np.nan
        # загружаем свой словарь обработки нестандартных адресов
    conv_df = pd.read_excel(r'..\conv_dictionary.xlsx', sheet_name='dict', engine='openpyxl')
    for con_quest in conv_df.parsed_name:
        if re.search(con_quest, line) is not None:
            print(con_quest)
            print(conv_df[conv_df.parsed_name == con_quest].correct_name)
            logging.info(str(con_quest))
            logging.info(str(conv_df[conv_df.parsed_name == con_quest].correct_name))
            line = line.replace(con_quest, conv_df[conv_df.parsed_name == con_quest].correct_name.tolist()[0])
    if re.search(r'\d+-[я] улица', line) is not None:
        st = re.search(r'\d+-[я] улица', line)
        st1 = st[0].split()
        line = re.sub(st[0], st1[1] + ' ' + st1[0], line)
    if re.search(r'\d+-[й] квартал', line) is not None:
        st = re.search(r'\d+-[й] квартал', line)
        st1 = st[0].split()
        line = re.sub(st[0], st1[1] + ' ' + st1[0], line)
    for _, paramValue in enumerate(addr_extractor.find(line).fact.parts):
        p0 = addr_extractor.find(line).fact.parts
    try:
        for p in p0:
            if p.type == 'город':
                if p.value != 'Москва' and p.value != 'Москвы':
                    adr_street += str(', г. ' + p.value)
            if p.type == 'деревня':
                adr_street += str(', д. ' + p.value)
            if p.type == 'микрорайон':
                adr_street += str(', мкр. ' + p.value)
            if p.type == 'бульвар':
                adr_street += str(', б-р. ' + p.value)
            if p.type == 'дачный поселок':
                adr_street += str(', дп. ' + p.value)
            if p.type == 'поселок':
                adr_street += str(', п. ' + p.value)
            if p.type == 'шоссе':
                adr_street += str(', ш. ' + p.value)
            if p.type == 'квартал':
                adr_street += str(', кв-л. ' + p.value)
            if p.type == 'улица':
                str_cons = p.value.split(' ')
                if str_cons[0] == '43':
                    p.value = '43-й Армии'
                if str_cons[-1] in ('Б', 'М', 'Нов', 'Стар', 'Нижн', 'Верхн', 'Ср'):
                    p.value = p.value + '.'
                if len(str_cons) > 1 and re.search(r'\d+-[я]', str_cons[0]) is not None:
                    adr = ''
                    for i in range(1, len(str_cons)):
                        adr += str_cons[i] + ' '
                    adr_street += str(', ул. ' + adr + str_cons[0])
                elif len(str_cons) > 1 and re.search('Больш', str_cons[0]) is not None:
                    adr = ''
                    for i in range(1, len(str_cons)):
                        adr += str_cons[i] + ' '
                    adr_street += str(', ул. ' + adr + 'Б.')
                elif len(str_cons) > 1 and re.search('Новая', str_cons[0]) is not None:
                    adr = ''
                    for i in range(1, len(str_cons)):
                        adr += str_cons[i] + ' '
                    adr_street += str(', ул. ' + adr + 'Нов.')
                elif len(str_cons) > 1 and re.search('Старая', str_cons[0]) is not None:
                    adr = ''
                    for i in range(1, len(str_cons)):
                        adr += str_cons[i] + ' '
                    adr_street += str(', ул. ' + adr + 'Стар.')
                elif len(str_cons) > 1 and re.search('Мал', str_cons[0]) is not None:
                    adr = ''
                    for i in range(1, len(str_cons)):
                        adr += str_cons[i] + ' '
                    adr_street += str(', ул. ' + adr + 'М.')
                elif len(str_cons) > 1 and re.search('Нижн', str_cons[0]) is not None:
                    adr = ''
                    for i in range(1, len(str_cons)):
                        adr += str_cons[i] + ' '
                    adr_street += str(', ул. ' + adr + 'Нижн.')
                elif len(str_cons) > 1 and re.search('Верхн', str_cons[0]) is not None:
                    adr = ''
                    for i in range(1, len(str_cons)):
                        adr += str_cons[i] + ' '
                    adr_street += str(', ул. ' + adr + 'Верхн.')
                else:
                    adr_street += str(', ул. ' + p.value)

            if p.type == 'переулок':
                str_cons = p.value.split(' ')
                if str_cons[-1] in ('Б', 'М', 'Нов', 'Стар', 'Нижн', 'Верхн', 'Ср'):
                    p.value = p.value + '.'
                if len(str_cons) > 1 and re.search(r'\d+-[й]', str_cons[0]) is not None:
                    adr = ''
                    for i in range(1, len(str_cons)):
                        adr += str_cons[i] + ' '
                    adr_street += str(', пер. ' + adr + str_cons[0])
                elif len(str_cons) > 1 and re.search('Средн', str_cons[0]) is not None:
                    adr = ''
                    for i in range(1, len(str_cons)):
                        adr += str_cons[i] + ' '
                    adr_street += str(', пер. ' + adr + 'Ср.')
                elif len(str_cons) > 1 and re.search('Малый', str_cons[0]) is not None:
                    adr = ''
                    for i in range(1, len(str_cons)):
                        adr += str_cons[i] + ' '
                    adr_street += str(', пер. ' + adr + 'М.')
                elif len(str_cons) > 1 and re.search('Большой', str_cons[0]) is not None:
                    adr = ''
                    for i in range(1, len(str_cons)):
                        adr += str_cons[i] + ' '
                    adr_street += str(', пер. ' + adr + 'Б.')
                else:
                    adr_street += str(', пер. ' + p.value)

            if p.type == 'проезд':
                str_cons = p.value.split(' ')
                if str_cons[-1] in ('Б', 'М', 'Нов', 'Стар', 'Нижн', 'Верхн', 'Ср'):
                    p.value = p.value + '.'
                if len(str_cons) > 1 and re.search(r'\d+-[й]', str_cons[0]) is not None:
                    adr = ''
                    for i in range(1, len(str_cons)):
                        adr += str_cons[i] + ' '
                    adr_street += str(', проезд. ' + adr + str_cons[0])
                elif len(str_cons) > 1 and re.search('Верхний', str_cons[0]) is not None:
                    adr = ''
                    for i in range(1, len(str_cons)):
                        adr += str_cons[i] + ' '
                    adr_street += str(', проезд. ' + adr + 'Верхн.')
                elif len(str_cons) > 1 and re.search('Больш', str_cons[0]) is not None:
                    adr = ''
                    for i in range(1, len(str_cons)):
                        adr += str_cons[i] + ' '
                    adr_street += str(', проезд. ' + adr + 'Б.')
                else:
                    adr_street += str(', проезд. ' + p.value)

            if p.type == 'проспект':
                adr_street += str(', пр-кт. ' + p.value)
            if p.type == 'аллея':
                adr_street += str(', аллея. ' + p.value)
            if p.type == 'площадь':
                str_cons = p.value.split(' ')
                if str_cons[-1] in ('Б', 'М', 'Нов', 'Стар', 'Нижн', 'Верхн', 'Ср'):
                    p.value = p.value + '.'
                if len(str_cons) > 1 and re.search('Большая', str_cons[0]) is not None:
                    adr = ''
                    for i in range(1, len(str_cons)):
                        adr += str_cons[i] + ' '
                    adr_street += str(', пл. ' + adr + 'Б.')
                else:
                    adr_street += str(', пл. ' + p.value)
            if p.type == 'набережная':
                adr_street += str(', наб. ' + p.value)
            if p.type == 'село':
                adr_street += str(', с. ' + p.value)
            if p.type == 'дом':
                adr_house = str(', д. ' + p.value)
            if p.type == 'корпус':
                adr_house = adr_house + str(', к. ' + p.value)
            if p.type == 'строение':
                adr_house = adr_house + str(', стр. ' + p.value)
            if p.type == 'этаж':
                adr_floor = p.value
        adr_street = re.sub('ё', 'е', adr_street)
        adr_street = 'г. Москва' + adr_street
        ret = [adr_street, adr_house, adr_floor]
    except:
        return ['не распознан', 'не распознан', 'не распознан']
    return ret


def secondary_addr_normalize_for_metrodistance(line):
    """
    Функция углубленной нормализации адреса. Получает на вход строку адреса и пытается ее распарсить,
    при этом проверяет вхождение слов из словаря conv_dictionary.xlsx и заменяет на соответствующие,
    переносит нумерацию улиц в конец строки (например, 1-я Трудовая должна быть "Трудовая 1-я ул."),
    обрабатывает названия Большой, Нижний и т.д., заменяет букву "ё" на "е".
    Так как делает больше проверок, запускается на части датафрейма, которую не удалось распарсить первично.
    Возвращает список из строки с указанием города/улицы и строки с указанием дома/корпуса/строения
    """
    # morph_vocab = MorphVocab()
    # addr_extractor = AddrExtractor(morph_vocab)
    adr_street = ''
    adr_house = ''
    if re.search(r'\d+-[я] улица', line) is not None:
        st = re.search(r'\d+-[я] улица', line)
        st1 = st[0].split()
        line = re.sub(st[0], st1[1] + ' ' + st1[0], line)
    if re.search(r'\d+-[й] квартал', line) is not None:
        st = re.search(r'\d+-[й] квартал', line)
        st1 = st[0].split()
        line = re.sub(st[0], st1[1] + ' ' + st1[0], line)
    for _, paramValue in enumerate(addr_extractor.find(line).fact.parts):
        p0 = addr_extractor.find(line).fact.parts
    for p in p0:
        if p.type == 'город':
            if p.value != 'Москва' and p.value != 'Москвы':
                adr_street += str(', г. ' + p.value)
        if p.type == 'деревня':
            adr_street += str(', д. ' + p.value)
        if p.type == 'микрорайон':
            adr_street += str(', мкр. ' + p.value)
        if p.type == 'бульвар':
            adr_street += str(', б-р. ' + p.value)
        if p.type == 'дачный поселок':
            adr_street += str(', дп. ' + p.value)
        if p.type == 'поселок':
            adr_street += str(', п. ' + p.value)
        if p.type == 'шоссе':
            adr_street += str(', ш. ' + p.value)
        if p.type == 'квартал':
            adr_street += str(', кв-л. ' + p.value)
        if p.type == 'улица':
            str_cons = p.value.split(' ')
            if str_cons[0] == '43':
                p.value = '43-й Армии'
            if len(str_cons) > 1 and re.search(r'\d+-[я]', str_cons[0]) is not None:
                adr = ''
                for i in range(1, len(str_cons)):
                    adr += str_cons[i] + ' '
                adr_street += str(', ул. ' + adr + str_cons[0])
            elif len(str_cons) > 1 and re.search('Больш', str_cons[0]) is not None:
                adr = ''
                for i in range(1, len(str_cons)):
                    adr += str_cons[i] + ' '
                adr_street += str(', ул. ' + adr + 'Б.')
            elif len(str_cons) > 1 and re.search('Мал', str_cons[0]) is not None:
                adr = ''
                for i in range(1, len(str_cons)):
                    adr += str_cons[i] + ' '
                adr_street += str(', ул. ' + adr + 'М.')
            elif len(str_cons) > 1 and re.search('Нижн', str_cons[0]) is not None:
                adr = ''
                for i in range(1, len(str_cons)):
                    adr += str_cons[i] + ' '
                adr_street += str(', ул. ' + adr + 'Нижн.')
            elif len(str_cons) > 1 and re.search('Верхн', str_cons[0]) is not None:
                adr = ''
                for i in range(1, len(str_cons)):
                    adr += str_cons[i] + ' '
                adr_street += str(', ул. ' + adr + 'Верхн.')
            else:
                adr_street += str(', ул. ' + p.value)
        if p.type == 'переулок':
            str_cons = p.value.split(' ')
            if len(str_cons) > 1 and re.search(r'\d+-[й]', str_cons[0]) is not None:
                adr = ''
                for i in range(1, len(str_cons)):
                    adr += str_cons[i] + ' '
                adr_street += str(', пер. ' + adr + str_cons[0])
            elif len(str_cons) > 1 and re.search('Средн', str_cons[0]) is not None:
                adr = ''
                for i in range(1, len(str_cons)):
                    adr += str_cons[i] + ' '
                adr_street += str(', пер. ' + adr + 'Ср.')
            else:
                adr_street += str(', пер. ' + p.value)
        if p.type == 'проезд':
            str_cons = p.value.split(' ')
            if len(str_cons) > 1 and re.search(r'\d+-[й]', str_cons[0]) is not None:
                adr = ''
                for i in range(1, len(str_cons)):
                    adr += str_cons[i] + ' '
                adr_street += str(', проезд. ' + adr + str_cons[0])
            elif len(str_cons) > 1 and re.search('Верхний', str_cons[0]) is not None:
                adr = ''
                for i in range(1, len(str_cons)):
                    adr += str_cons[i] + ' '
                adr_street += str(', проезд. ' + adr + 'Верхн.')
            elif len(str_cons) > 1 and re.search('Больш', str_cons[0]) is not None:
                adr = ''
                for i in range(1, len(str_cons)):
                    adr += str_cons[i] + ' '
                adr_street += str(', проезд. ' + adr + 'Б.')
            else:
                adr_street += str(', проезд. ' + p.value)
        if p.type == 'проспект':
            adr_street += str(', пр-кт. ' + p.value)
        if p.type == 'аллея':
            adr_street += str(', аллея. ' + p.value)
        if p.type == 'площадь':
            adr_street += str(', пл. ' + p.value)
        if p.type == 'набережная':
            adr_street += str(', наб. ' + p.value)
        if p.type == 'село':
            adr_street += str(', с. ' + p.value)
        if p.type == 'дом':
            adr_house = str(', д. ' + p.value)
        if p.type == 'корпус':
            adr_house = adr_house + str(', к. ' + p.value)
        if p.type == 'строение':
            adr_house = adr_house + str(', стр. ' + p.value)
    adr_street = re.sub('ё', 'е', adr_street)
    adr_street = 'г. Москва' + adr_street
    ret = adr_street + adr_house
    return ret


def clean_gkh_add(path: str, sheet_name: str = "Sheet1"):
    """
    Функция очистки данных в конкретном листе.
    :param path: Путь до файла Excel
    :param sheet_name: Имя листа в таблице Excel, куда буду писать данные
    :return: возвращает файл пустой файл с названиями столбцов
    """
    wb = ox.load_workbook(path)
    wb[sheet_name].delete_rows(2, 1000)
    wb.save(path)


def add_unrecognized():
    """ Функция добавляет данные из файла new_buildings.xlsx в базу ЖКХ, сохраняет в csv всю базу,
        удаляет импортированные записи из файла new_buildings.xlsx.
    """
    global gkh_df
    print('Добавляются данные в базу ЖКХ')
    logging.info('Добавляются данные в базу ЖКХ')
    gkh_add = pd.read_excel(r'..\new_buildings.xlsx', sheet_name='Sheet1', engine='openpyxl', keep_default_na=False)
    gkh_df = gkh_df.append(gkh_add, ignore_index=True)
    gkh_df = gkh_df.drop_duplicates(subset=['address_w'], keep='last')
    gkh_df.to_csv('buildings_oper.csv', index=False)
    clean_gkh_add(r'..\new_buildings.xlsx', sheet_name='Sheet1')


def update_gkh_base(address, res_str):
    """ обновление информации по адресу в базе ЖКХ.
        Не забыть скачать базу до и сохранить в файл после работы  """
    gkh_df.loc[gkh_df.address_w == address, 'num_floors'] = res_str[0]
    if res_str[1] != '':
        gkh_df.loc[gkh_df.address_w == address, 'metro_name_gkh'] = res_str[1]
    else:
        gkh_df.loc[gkh_df.address_w == address, 'metro_name_gkh'] = np.nan
    if res_str[3] != '':
        gkh_df.loc[gkh_df.address_w == address, 'metro_minutes'] = res_str[3]
    else:
        gkh_df.loc[gkh_df.address_w == address, 'metro_minutes'] = np.nan
    gkh_df.loc[gkh_df.address_w == address, 'metro_km'] = res_str[2]
    try:
        gkh_df.loc[gkh_df.address_w == address, 'year_exp_w'] = int(res_str[4])
    except (Exception, ):
        print('год не удалось сохранить')
        logging.info('год не удалось сохранить')


def update_new_buildings(path: str, _df, startcol: int = 1, startrow: int = 1, sheet_name: str = "Sheet1"):
    """Функция добавления данных в new_buildings.
    :param path: Путь до файла Excel
    :param _df: Датафрейм Pandas для записи
    :param startcol: Стартовая колонка в таблице листа Excel, куда буду писать данные
    :param startrow: Стартовая строка в таблице листа Excel, куда буду писать данные
    :param sheet_name: Имя листа в таблице Excel, куда буду писать данные
    :return: возвращает файл с датафреймом  (названия столбцов предустановлены),
    """
    wb = ox.load_workbook(path)
    for ir in range(0, len(_df)):
        for ic in range(0, len(_df.iloc[ir])):
            wb[sheet_name].cell(startrow + ir, startcol + ic).value = _df.iloc[ir][ic]
    wb.save(path)


def metro_and_floor_data(address_w):
    """ Функция поиска данных об адресе на ресурсе flatInfo.ru (этажность и расстояние до метро)
    Функция дополнительно обновляет базу ЖКХ update_gkh_base    """
    global gkh_df

    print(address_w)
    logging.info(str(address_w))
    res = [np.nan for _ in range(5)]
    try:
        driver = start_browser_for_parse()
        driver.get('https://flatinfo.ru')
        element = driver.find_element(By.XPATH, "//div[@class='search-home input-group search-home_show']/input[1]")
        adr1 = re.sub('проезд.', 'проезд', address_w)
        adr1 = re.sub('пр-кт.', 'проспект', adr1)
        element.send_keys(adr1)
        time.sleep(3)
        butt = driver.find_element(By.XPATH, "//div[@class='search-home input-group search-home_show']/button[1]")
        ActionChains(driver).click(butt).perform()
        time.sleep(3)
        try:
            ActionChains(driver).click(butt).perform()
        except (Exception, ):
            pass

        # и сохраняем адрес перехода
        cur_url = driver.current_url
        print(cur_url)
        logging.info(str(cur_url))
        # если переход произошел на страницу с данными о доме
        if re.search('h_info', cur_url) is not None:
            addr_line = re.sub(' в Москве', '',
                               driver.find_element(By.XPATH, "//h1[starts-with(text(), 'О доме')]").text)
            parse_addr_line = secondary_addr_normalize_for_metrodistance(addr_line)
            if (parse_addr_line.lower()) != address_w:
                print('адреса не совпадают')
                logging.info('адреса не совпадают')
                driver.quit()
                return res
            print('адреса норм')
            logging.info('адреса норм')
            try:
                r = requests.get(cur_url)
                soup = BeautifulSoup(r.text, 'lxml')
                # скачиваем страницу инфы о доме
                full_info = soup.find('div', class_='page__content')
                li_blocks = full_info.find_all('li', class_='fi-list__item fi-list-item')
                # и ищем характеристики этажности
                for li_bl in li_blocks:
#                    q = li_bl.text.strip('\n').split(':')
#                     if q[0] == 'Этажей всего':
#                         res[0] = q[1].strip('\n').split('\n')[0]
#                     if q[0] == 'Год постройки':
#                         res[4] = q[1].strip('\n').strip()
                    if li_bl.find('span', class_='fi-list-item__label').text == 'Этажей всего':
                        t1 = li_bl.find('span', class_='fi-list-item__value').text
                        t2 = t1.strip('\r').strip('\n').strip('\t').split('\n')
                        t3 = t2[0].strip()
                        res[0] = t3
                        # res[0] = li_bl.find('span', class_='fi-list-item__value').text\
                        #                 .strip('\n').split('\n')[0].strip()
                    if li_bl.find('span', class_='fi-list-item__label').text == 'Год постройки':
                        res[4] = li_bl.find('span', class_='fi-list-item__value').text\
                                        .strip('\n').strip()

            except (Exception, ):
                print('данные о доме не получены')
                logging.info('данные о доме не получены')

            cur_url = cur_url.replace('info1', 'info2')
            r = requests.get(cur_url)
            soup = BeautifulSoup(r.text, 'lxml')
            try:
                full_info = soup.find_all('div', class_='col-md-6')
                metro_info = ['нет данных', 'нет данных', 'нет данных']
                for bl in full_info:
                    if bl.find('h2').text == 'Метро рядом':
                        metro_bl = bl.find_all('td')
                        metro_info = []
                        for block in metro_bl[:3]:
                            metro_info.append(block.text)
                res[1] = metro_info[0]
                if metro_info[1] != 'нет данных':
                    res[2] = round(float(metro_info[1].split()[0])/1000, 2)
                if metro_info[2] != 'нет данных':
                    res[3] = int(metro_info[2].split()[0])
            except (Exception, ):
                pass
        driver.quit()
    except (Exception, ):
        pass
    try:
        driver.quit()
    except (Exception, ):
        pass
    update_gkh_base(address_w, res)
    print(res)
    logging.info(f'Объект {res}')
    if sum(pd.isna(res)) == 5:
        print('данные об объекте', address_w, ' не получены')
        logging.info('данные об объекте' + str(address_w) + ' не получены')
    if 0 < sum(pd.isna(res)) < 5:
        print('данные об объекте', address_w, ' не полные')
        logging.info('данные об объекте' + str(address_w) + ' не полные')
    return res


def renov_fill(ser):
    if str(ser[0]) != 'nan':
        return ser[0]
    elif ser[1] >= 2010:
        return 'новый дом'
    elif ser[1] < 2010:
        return 'нет в плане'
    else:
        return 'N/A'


def max_floor(st):
    """ Функция очистки данных об этажности. Возвращает последнее число из строки и переводит в int"""
    try:
        return int(re.findall(r'\d+', st)[-1])
    except (Exception, ):
        return 0


def qty_rooms_predict(ser):
    """ Функция предсказывает количествокомнат исходя из данных переданной серии pd.Series
        на обученном классификаторе дерева решений clf
        Возвращает серию"""
    global clf

    if pd.isna(ser.loc['qty_rooms']):
        try:
            fea_ = np.array(ser[['total_floor_ml', 'year_exp_w', 'obj_square']]).reshape(1, -1)
            ser.loc['qty_rooms'] = clf.predict(fea_)[0]
            ser.loc['is_qty_rooms_predicted'] = 1
        except Exception as e:
            ser.loc['qty_rooms'] = 0
            ser.loc['is_qty_rooms_predicted'] = 0
    else:
        ser.loc['is_qty_rooms_predicted'] = 0
    return ser


def update_spreadsheet_hist(path: str, _df, startcol: int = 1, startrow: int = 1, sheet_name: str = "Прошедшие"):
    """Функция перезаписи данных в конкретный лист "Прошедшие".
    :param path: Путь до файла Excel
    :param _df: Датафрейм Pandas для записи
    :param startcol: Стартовая колонка в таблице листа Excel, куда буду писать данные
    :param startrow: Стартовая строка в таблице листа Excel, куда буду писать данные
    :param sheet_name: Имя листа в таблице Excel, куда буду писать данные
    :return: возвращает файл с датафреймом data_for_save (названия столбцов предустановлены),
            в ячейку А1 пишется дата создания файла"""
    global date_to
    thin_border = Border(left=Side(style='thin'),
                         right=Side(style='thin'),
                         top=Side(style='thin'),
                         bottom=Side(style='thin'))
    wb = ox.load_workbook(path)
    for ir in range(0, len(_df)):
        for ic in range(0, len(_df.iloc[ir])):
            wb[sheet_name].cell(startrow + ir, startcol + ic).value = _df.iloc[ir][ic]
            wb[sheet_name].cell(startrow + ir, startcol + ic).border = thin_border
            if ic in [5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 24]:
                wb[sheet_name].cell(startrow + ir, startcol + ic).alignment = Alignment(horizontal='center')
            elif ic in [15, 16, 17, 18]:
                wb[sheet_name].cell(startrow + ir, startcol + ic).alignment = Alignment(horizontal='center')
                wb[sheet_name].cell(startrow + ir, startcol + ic).number_format = r'# ##0'
            elif ic in [21]:
                wb[sheet_name].cell(startrow + ir, startcol + ic).alignment = Alignment(horizontal='center')
                wb[sheet_name].cell(startrow + ir, startcol + ic).number_format = 'DD.MM.YYYY'
            elif ic in [19]:
                wb[sheet_name].cell(startrow + ir, startcol + ic).alignment = Alignment(horizontal='center')
                wb[sheet_name].cell(startrow + ir, startcol + ic).number_format = '0.00%'
            elif ic in [25]:
                wb[sheet_name].cell(startrow + ir, startcol + ic).alignment = Alignment(horizontal='center')
                wb[sheet_name].cell(startrow + ir, startcol + ic).font = Font(color='808080')
                wb[sheet_name].cell(startrow + ir, startcol + ic).fill = PatternFill('solid', fgColor="D9D9D9")

    wb[sheet_name].cell(1, 1).value = pd.Timestamp('today').date() - pd.Timedelta(1, "d")
    wb[sheet_name].conditional_formatting = ConditionalFormattingList()
    cells_range = 'N2:N' + str(len(_df) + 1)
    pred_font = Font(color='0000FF')
    diff_style = DifferentialStyle(font=pred_font)
    rule = Rule(type="expression", dxf=diff_style)
    rule.formula = ["O2=1"]
    wb[sheet_name].conditional_formatting.add(cells_range, rule)
    wb.save(path)


def start_browser_for_parse():
    """ Функция запуска Chrome-браузера"""
    chrome_options = webdriver.ChromeOptions()
    # chrome_options.add_argument("--headless")
    # chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument('--start-maximized')
    chrome_options.add_argument('--log-level=3')
    chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
    prefs = {"profile.managed_default_content_settings.images": 2}
    chrome_options.add_experimental_option("prefs", prefs)
    path_to_wd = os.getcwd()
    driver = webdriver.Chrome(executable_path=(path_to_wd + '\\chromedriver.exe'), options=chrome_options)
    return driver


def get_url_list_from_page(driver, cur_page_url):
    """ Функция получения списка URL всех лотов на странице cur_page_url
        При открытии страницы разворачивает все svg-объекты.
    :param driver: работающий в программе WebDriver
    :param cur_page_url: URL страницы investmoscow.ru для получения списка лотов
    :return: возвращает список URL
    """
    url_page_lst = []
    driver.get(cur_page_url)
    wait = WebDriverWait(driver, 20)
    a_class = (By.CLASS_NAME, "uid-mb-40")
    qty_a_class = len(driver.find_elements_by_class_name("uid-mb-40"))
    if qty_a_class > 0:
        wait.until(
            EC.presence_of_element_located(a_class))
    # ---> прокрутка до конца страницы
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    driver.execute_script("window.scrollTo(document.body.scrollHeight, 0);")
    # ---<
    time.sleep(2)
    arrows_down_qty = len(driver.find_elements(By.XPATH,
                                               '//div[@class="list"]/div/div[@class="uid-mb-40"]' +
                                               '/div[@class="uid-group-card__collapse"]/div[@class="uid-group-card-toggle"]/' +
                                               '/div/*[name()="svg"][@class="uid-arrow uid-arrow-down"]'))
    while arrows_down_qty > 0:
        #        print(arrows_down_qty)
        try:
            # wait.until(EC.presence_of_element_located((By.XPATH,
            driver.find_element(By.XPATH,
                                '//div[@class="list"]/div/div[@class="uid-mb-40"]' +
                                '/div[@class="uid-group-card__collapse"]/div[@class="uid-group-card-toggle"]/' +
                                '/div/*[name()="svg"][@class="uid-arrow uid-arrow-down"]').click()
            arrows_down_qty = len(driver.find_elements(By.XPATH,
                                                       '//div[@class="list"]/div/div[@class="uid-mb-40"]' +
                                                       '/div[@class="uid-group-card__collapse"]/div[@class="uid-group-card-toggle"]/' +
                                                       '/div/*[name()="svg"][@class="uid-arrow uid-arrow-down"]'))
        except (Exception,) as exc:
            pass

    # ------------------   конец - корректировка поиска стрелочек вниз
    wait.until(
        EC.presence_of_element_located((By.XPATH,
                                        '//div[@class="list"]')))
    body = driver.find_element_by_css_selector('body')

    # ----> медленная прокрутка окна
    driver.execute_script("window.scrollTo(document.body.scrollHeight, 0);")
    while True:
        # Scroll down to bottom
        body.send_keys(Keys.PAGE_DOWN)
        scroll_position = driver.execute_script("return window.scrollY;")
        max_position = driver.execute_script("return document.body.scrollHeight;")
        # Wait to load page
        time.sleep(SCROLL_PAUSE_TIME)
        # Calculate new scroll height and compare with last scroll height
        if scroll_position >= max_position - 979:
            print("break")
            break
    soup = BeautifulSoup(driver.find_element(By.XPATH, '//div[@class="list"]').get_attribute('innerHTML'),
                         "html.parser")
    nfo = soup.find_all('a', attrs={"class": "uid-tenders-card__main"})
    for el in nfo:
        url_page_lst.append('https://investmoscow.ru' + el['href'])
    return driver, url_page_lst


def reparsing_unreadable_urls(driver, lot_df, qty_lots, unreadable_urls):
    """ Функция повторного парсинга непрочитанных ранее лотов.
        :param lot_df: датафрейм с уже обработанными лотами
        :param qty_lots: количество всех лотов
        :param unreadable_urls: список URL необработанных лотов
        :return: датафрейм с обработанными лотами. Если в итоге не получается скачать некоторые лоты -
        выводит данные об этом.
        """
    for j in range(10):
        oper_lst = unreadable_urls.copy()
        if len(oper_lst) > 0:
            print(f'Попытка чтения списка номер {j + 2}')
            print(f'Нераспознанных ссылок {len(unreadable_urls)}')
            logging.info(f'Попытка чтения списка номер {j + 2}')
            logging.info(f'Нераспознанных ссылок {len(unreadable_urls)}')
            for url in tqdm(oper_lst):
                # считываем страницу с заданным URL
                try:
                    # page = requests.get(url, timeout=60)
                    # soup = BeautifulSoup(page.text, "html.parser")
                    # lot_single = parse_lot(soup, is_actual=is_actual_lots)
                    driver, lot_single = parse_lot(driver, url)
                    lot_single['investMSK_URL'] = url
                    lot_df = pd.concat([lot_df, lot_single.to_frame().T], ignore_index=True)
                    unreadable_urls.remove(url)
                except (Exception,):
                    print(f'Нечитаемая ссылка {url}')
                    logging.info(f'Нечитаемая ссылка {url}')
                    continue
        else:
            break
    if len(lot_df) == qty_lots:
        print('Все объекты скачаны')
        logging.info('Все объекты скачаны')
    else:
        print('Не удалось скачать объектов: ', len(unreadable_urls))
        logging.info(f'Не удалось скачать объектов: {len(unreadable_urls)}')
        print(unreadable_urls)
        logging.info(f'Объекты {unreadable_urls}')
    return lot_df


def drop_duplicated_lots(lot_df):
    """ Функция чистки датафрейма от дублирующих лотов по lot_tag. (Остаются последние по дате аукциона)
        :param lot_df: полный датафрейм после парсинга лотов
        :return: датафрейм с уникальными лотами. Выводит данные о количестве уникальных лотов.
    """
    # поиск дублирующихся лотов
    last_dup = lot_df[lot_df.duplicated(['lot_tag'], keep=False)].sort_values('auct_date')
    last_dup = last_dup.drop_duplicates(subset=['lot_tag'], keep='last')
    # удаляем дублирующиеся лоты по тагу
    lot_df = lot_df.drop_duplicates(subset=['lot_tag'], keep=False, ignore_index=True)
    # и дописываем оптимальные из дубликатов к датафрейму
    lot_df = lot_df.append(last_dup).reset_index(drop=True)
    print('Количество уникальных объектов: ', len(lot_df))
    logging.info(f'Количество уникальных объектов: {len(lot_df)}')
    return lot_df


def compare_with_final_df(lot_df, final_df):
    lot_df = lot_df.merge(final_df[['lot_tag', 'addr_norm']].rename(columns={'addr_norm':'is_present'}),
                          how='left', on='lot_tag')
    lot_df = lot_df[lot_df.is_present.isna()]
    lot_df = lot_df.drop(columns=['is_present'])
    return lot_df


def recognize_and_normalize_addresses(norm_df):
    """ Функция нормализации адресов датафрейма (приведения к единому стандарту).
        Сравниваются адреса, полученные с помощью NER-модуля Natasha из объявлений, с адресами в базе ЖКХ.
        При отсутствии совпадений - провдоится повтоная нормализация с использованием словаря преобразований
        conv_dictionary.xlsx. Если совпадений нет и адрес распознан нормально (определяется путем сверки
        строки адреса с преобразованным адресом визуально) - можно добавить данные вручную в базу ЖКХ.
        :param norm_df: полный датафрейм после парсинга лотов
        :return: norm_df: датафрейм с базой ЖКХ и нормализованными адресами.
    """
    global gkh_df
    # считываем базу адресов с указанием года постройки и ссылки на дом (далее - база ЖКХ)
    print('Считывается база ЖКХ')
    logging.info('Считывается база ЖКХ')
    gkh_df = pd.read_csv('buildings_oper.csv')
    gkh_df['address_w'] = gkh_df.address_w.apply(lambda x: str(x).lower())

    # проводим первичную нормализацию адресов датафрейма
    print('Первичная нормализация адресов датафрейма')
    logging.info('Первичная нормализация адресов датафрейма')
    norm_df.loc[:, 'addr_street'], norm_df.loc[:, 'addr_build_num'], norm_df.loc[:, 'addr_floor_from_title'] = zip(
        *norm_df.loc[:, 'addr_string'].progress_apply(primary_addr_normalize))

    # добавляем колонки с нормализованным адресом, и приведенным к нижнему регистру
    # по приведенному к нижнему регистру адресу пытаемся сджойнить с базой ЖКХ.
    norm_df['addr_norm'] = norm_df['addr_street'] + norm_df['addr_build_num']
    norm_df['addr_norm_lower'] = norm_df['addr_norm'].apply(lambda x: str(x).lower())
    norm_df = norm_df.merge(gkh_df, how='left', left_on='addr_norm_lower', right_on='address_w')

    # выделяем из базы строки, для которых не нашлось соответствия в базе ЖКХ
    quest_df = norm_df[norm_df.address_w.isna()]
    norm_df = norm_df.drop(norm_df[norm_df.address_w.isna()].index)

    # создаем цикл обработки необработанных адресов
    if len(quest_df) > 0:
        loop_flag = True
    else:
        loop_flag = False
    while loop_flag:
        # если обновлялась база ЖКХ
        # считываем базу адресов с указанием года постройки и ссылки на дом (далее - база ЖКХ)
        print('Повторная нормализация необработанных ранее адресов датафрейма')
        logging.info('Повторная нормализация необработанных ранее адресов датафрейма')
        print('Считывается база ЖКХ')
        gkh_df = pd.read_csv('buildings_oper.csv')
        gkh_df['address_w'] = gkh_df.address_w.apply(lambda x: str(x).lower())
        # убираем из нее ранее добавленные данные из ЖКХ (чтобы потом повторно попробовать сджойнить)
        quest_df = quest_df.drop(
            columns=['address_w', 'adm_area', 'mun_district', 'year_exp_w', 'building_page', 'metro_minutes',
                     'num_floors', 'metro_name_gkh', 'metro_km'])
        # проводим повторную нормализацию адресов датафрейма
        quest_df.loc[:, 'addr_street'], quest_df.loc[:, 'addr_build_num'], quest_df.loc[:, 'addr_floor_from_title'] \
            = zip(*quest_df.loc[:, 'addr_string'].progress_apply(secondary_addr_normalize))
        # добавляем колонки с нормализованным адресом, и приведенным к нижнему регистру
        # по приведенному к нижнему регистру адресу пытаемся сджойнить с базой ЖКХ.
        quest_df['addr_norm'] = quest_df['addr_street'] + quest_df['addr_build_num']
        quest_df['addr_norm_lower'] = quest_df['addr_norm'].apply(lambda x: str(x).lower())
        quest_df = quest_df.merge(gkh_df, how='left', left_on='addr_norm_lower', right_on='address_w')
        # те строки, которые получилось сджойнить, добавляем в norm_df
        norm_df = norm_df.append(quest_df.query('addr_norm_lower == address_w'))
        quest_df = quest_df.query('addr_norm_lower != address_w')
        if len(quest_df) == 0:
            break
        else:
            print('Нераспознанных адресов: ', len(quest_df))
            logging.info(f'Нераспознанных адресов: {len(quest_df)}')
            # новая версия - выдает на печать только по одной паре реальный адрес / распознанный
            # quest_df = quest_df.reset_index(drop=True)
            quest_df = quest_df.reset_index(drop=True)
            quest_df_sample = quest_df.drop_duplicates(subset=['addr_norm_lower']).copy()
            for i in range(len(quest_df_sample)):
                try:
                    print(quest_df_sample.iloc[i,[7, 17]].values)
                    logging.info(f'{quest_df_sample.iloc[i,[7, 17]]}.values')
                except Exception as e:
                    logging.info(f'Ошибка: {traceback.format_exc()}')

            print('Если нормализованный адрес совпадает с адресом в строке объявления - ')
            print('заполните форму в файле new_buildings.xlsx. !!!! НЕ ЗАБУДЬТЕ ЗАКРЫТЬ ФАЙЛ')
            print('(после добавления данных в основную базу содержимое файла будет удалено)')
            print()
            print('Если нормализованный адрес распознан неверно - можно попытаться преобразовать его')
            print('для этого запустите файл addr_test.exe, скопируйте туда строку нераспознанного адреса')
            print('Если получилось преобразовать ее для нормального распознавания - внесите изменения в conv_dict.xlsx')
            print()
            print()
            print('После внесения изменений в данные            - нажмите \'Y\'')
            print('или продолжить без данных по этим адресам    - нажмите \'N\'')
            is_cont = input()
            while is_cont not in ['Y', 'y', 'N', 'n']:
                print('Некорректный ввод. Повторите выбор (Y/N)')
                is_cont = input()
            if is_cont in ['Y', 'y']:
                add_unrecognized()
                quest_df = quest_df.drop(
                    columns=['address_w', 'adm_area', 'mun_district', 'year_exp_w', 'building_page', 'metro_minutes',
                             'num_floors', 'metro_name_gkh', 'metro_km', 'address_w'])
                quest_df['addr_norm_lower'] = quest_df['addr_norm'].apply(lambda x: str(x).lower())
                quest_df = quest_df.merge(gkh_df, how='left', left_on='addr_norm_lower', right_on='address_w')
                norm_df = norm_df.append(quest_df.query('addr_norm_lower == address_w'))
                quest_df = quest_df.query('addr_norm_lower != address_w')
                if len(quest_df) == 0:
                    print('Все адреса распознаны')
                    logging.info('Все адреса распознаны')
                    loop_flag = False
                else:
                    print('Продолжить попытки распознавания адресов? (Y/N')
                    is_cont = input()
                    while is_cont not in ['Y', 'y', 'N', 'n']:
                        print('Некорректный ввод. Повторите выбор (Y/N)')
                        is_cont = input()
                    if is_cont in ['N', 'n']:
                        loop_flag = False
            else:
                loop_flag = False
    return norm_df


def is_metro_and_floor_data_complete(norm_df):
    """ Функция проверки полноты данных о привязке к метро и этажности зданий. При отсутствии данных в базе ЖКХ -
        запускается попытка скачать с сайта FlatInfo.ru, при невозможности - дает возможность доплнения данных
        базы ЖКХ вручную (через new_buildings.xlsx)
        :param norm_df: полный датафрейм после парсинга лотов
        :return: norm_df: датафрейм с базой ЖКХ и нормализованными адресами.
    """
    print('Проверка наличия данных о метро и этажности зданий')
    logging.info('Проверка наличия данных о метро и этажности зданий')
    out_of_data = norm_df[norm_df.metro_name_gkh.isna() | norm_df.num_floors.isna() | norm_df.year_exp_w.isna()] \
        .drop_duplicates(subset=['addr_norm'], keep='last')
    if out_of_data.shape[0] != 0:
        print('Отсутвуют данные по ', len(out_of_data), ' объектам')
        logging.info(f'Отсутвуют данные по {len(out_of_data)} объектам')
        print('Дождитесь окончания сбора и сохранения данных')
        out_of_data.progress_apply(lambda x: metro_and_floor_data(x.address_w), axis=1)
        gkh_df.to_csv('buildings_oper.csv', index=False)
        print('Данные сохранены')
        logging.info('Данные сохранены')
        norm_df = norm_df.drop(
            columns=['address_w', 'adm_area', 'mun_district', 'year_exp_w', 'building_page', 'metro_minutes',
                     'num_floors', 'metro_name_gkh', 'metro_km'])
        norm_df = norm_df.merge(gkh_df, how='left', left_on='addr_norm_lower', right_on='address_w')
        out_of_data = norm_df[norm_df.metro_name_gkh.isna() | norm_df.num_floors.isna()] \
            .drop_duplicates(subset=['addr_norm'], keep='last')
    if out_of_data.shape[0] != 0:
        update_new_buildings(
                            r'..\new_buildings.xlsx',
                            out_of_data[['adm_area', 'mun_district', 'address_w', 'year_exp_w', 'building_page',
                                        'num_floors', 'metro_name_gkh', 'metro_minutes', 'metro_km']],
                            startcol=1, startrow=2)
        print('Не все данные распознались успешно')
        logging.info('Не все данные распознались успешно')
        print('Внесите данные о метро и этажности вручную в файл new_buildings.xlsx.')
        print('Затем сохраните и закройте файл. Нажмите Enter.')
        input()
        # обновляем базу адресов с указанием года постройки и ссылки на дом (далее - база ЖКХ)
        add_unrecognized()
        norm_df = norm_df.drop(
            columns=['address_w', 'adm_area', 'mun_district', 'year_exp_w', 'building_page', 'metro_minutes',
                     'num_floors', 'metro_name_gkh', 'metro_km'])
        norm_df = norm_df.merge(gkh_df, how='left', left_on='addr_norm_lower', right_on='address_w')
        out_of_data = norm_df[norm_df.metro_name_gkh.isna() | norm_df.num_floors.isna()] \
            .drop_duplicates(subset=['addr_norm'], keep='last')
    if out_of_data.shape[0] == 0:
        print('Все данные об этажности и метро распознаны')
        logging.info('Все данные об этажности и метро распознаны')
    else:
        print('Остались неуточненные данные')
        print(out_of_data.addr_norm)
        logging.info('Остались неуточненные данные')
        logging.info(f'{out_of_data.addr_norm}')
    return norm_df


def fill_spaces_in_data(norm_df):
    """ Функция заполняет пропуски в данных, добавляет информацией о реновации,
        добавляет timestamp для сортировки.
        :param norm_df: полный датафрейм после парсинга лотов
        :return: norm_df: датафрейм с базой ЖКХ и нормализованными адресами.
    """
    # заполняем пропуски данными из подтянутых баз
    norm_df = norm_df.reset_index(drop=True)
    norm_df['total_floors'] = norm_df['total_floors'].fillna(norm_df['num_floors'])
    norm_df['addr_floor'] = norm_df['addr_floor'].fillna(norm_df['addr_floor_from_title'])
    norm_df['total_floors'] = norm_df['total_floors'].fillna(norm_df['num_floors'])
    norm_df['addr_floor'] = norm_df['addr_floor'].fillna(norm_df['addr_floor_from_title'])
    # добавляем колонку timestamp от дедлайна (для сортировки в Экселе)
    norm_df.auct_date = pd.to_datetime(norm_df.auct_date)
    norm_df['timestamp'] = norm_df['auct_date'].values.astype(np.int64) // 10 ** 9
    # добавляем (пока пустые) колонки "ремонт"
    norm_df['ремонт'] = ''
    # ---- добавляем реновацию
    renov_df = pd.read_csv('renovation_data.csv')
    norm_df = norm_df.merge(renov_df, how='left', on='addr_norm_lower')
    norm_df['year_exp_w'] = pd.to_numeric(norm_df['year_exp_w'], errors='coerce')
    norm_df['реновация'] = norm_df[['период реновации', 'year_exp_w']].apply(lambda x: renov_fill(x), axis=1)
    return norm_df


def fill_qty_rooms_with_predictions(norm_df):
    """ Функция заполняет пропуски в данных о количестве комнат с ипользованием ML-модели (DecisionTree),
        обучаемой на датасете прошедших торгов с известными параметрами
        (кол-во комнат <--- ujl постройки дома, этажность, площадь объекта).
        :param norm_df: полный датафрейм после парсинга лотов
        :return: norm_df: датафрейм с базой ЖКХ и нормализованными адресами.
    """
    global clf
    # скачиваем исторические данные
    df1 = pd.read_csv('historical_data.csv')
    df1 = df1[~df1.qty_rooms.isna()]
    df1 = df1[df1.is_qty_rooms_predicted == 0]
    df1 = df1[['total_floors', 'year_exp_w', 'qty_rooms', 'obj_square']]
    df1.total_floors = df1.total_floors.apply(lambda x: max_floor(str(x)))
    # и обучаем на них дерево решений
    clf = tree.DecisionTreeClassifier(criterion='entropy', max_depth=10)
    X = df1[['total_floors', 'year_exp_w', 'obj_square']]
    y = df1[['qty_rooms']]
    clf.fit(X, y)
    # чистим данные об этажности
    norm_df['total_floor_ml'] = norm_df.total_floors.apply(lambda x: max_floor(str(x)))
    # и пронозируем количество комнат
    norm_df = norm_df.apply(lambda x: qty_rooms_predict(x), axis=1)
    return norm_df


def compare_with_existing(ser):
    """ Функция сравнения Series, переданной из нового датафрейма, со строками в скачанном файле
        MosInvestData.xlsx.
        Пытается найти номер переданного лота в существующем датафрейме.
        При отстутсвии возвращает -1 в поле flag.
        При наличии в датафрейме - сравнивает данные, отличающиеся данные заменяются на новые,
        в поле "Примечания" вписывается дата внесения изменений, flag = -2
        Если данные корректны - возвращает номер строки датафрейма
    """
    global MID_df
    today = dt.datetime.today().date()
    try:
        row = MID_df.loc[MID_df.lot_tag == ser.lot_tag].index[0]
    except (Exception, ):
        return -1
    if str(MID_df.loc[row, 'deadline']) != str(ser.loc['deadline']):
        MID_df.loc[row, 'deadline'] = ser.loc['deadline']
        MID_df.loc[row, 'Примечания'] = today
        return -2
    if str(MID_df.loc[row, 'auct_date']) != str(ser.loc['auct_date']):
        MID_df.loc[row, 'auct_date'] = ser.loc['auct_date']
        MID_df.loc[row, 'Примечания'] = today
        return -2
    return row


def metro_prep(name):
    """ Функция перевода имени метро в нижний регистр и уточнения отдельных станций"""
    try:
        if name.lower() == 'покровская':
            name = 'покровское'
        elif name.lower() == 'кубанская (люблино)':
            name = 'люблино'
        elif name.lower() == 'новые черёмушки':
            name = 'новые черемушки'
        name = name.lower()
    except (Exception,):
        pass
    return name


def copy_cell(src_sheet, src_row, src_col,
              tgt_sheet, tgt_row, tgt_col,
              copy_style=True):
    """Функция копирования свойств ячейки с одного листа на другой вместе с содержимым """
    cell = src_sheet.cell(src_row, src_col)
    new_cell = tgt_sheet.cell(tgt_row, tgt_col, cell.value)
    if cell.has_style and copy_style:
        new_cell._style = copy(cell._style)


def update_spreadsheet_w_drop(path: str, _df, startcol: int = 1, startrow: int = 1, sheet_name: str = "Sheet1",
                              add_rows_qty: int = 0):
    """
    Функция перезаписи данных в конкретный лист с формартированием - применимо для шаблонов листов "Активные"
    и "NEW!!!". Копирует в лист "избранное" строки, помеченные заливкой на листе "Активные", подлежащие удалению
    в связи с истечением срока (дата аукциона+2дня -> flag = -2). Устанавливает правила условного форматирования
    данных. Если в строке flag=False (поле 30) --> строка удаляется
    :param path: Путь до файла Excel
    :param _df: Датафрейм Pandas для записи
    :param startcol: Стартовая колонка в таблице листа Excel, куда буду писать данные
    :param startrow: Стартовая строка в таблице листа Excel, куда буду писать данные
    :param sheet_name: Имя листа в таблице Excel, куда буду писать данные
    :param add_rows_qty: Количество строк для добавления выше последней
    :return: возвращает файл с датафреймом _df (названия столбцов предустановлены),
            в ячейку А1 пишется дата создания файла

    """
    thin_border = Border(left=Side(style='thin'),
                         right=Side(style='thin'),
                         top=Side(style='thin'),
                         bottom=Side(style='thin'))
    today = pd.Timestamp('today').date()
    file_name = path
    wb = ox.load_workbook(path)
    new_row_fav = wb['Избранное'].max_row + 1
    cnt_fav = 0
    if add_rows_qty > 0:
        wb[sheet_name].insert_rows(len(_df) + startrow - add_rows_qty, add_rows_qty)
    for r_ in range(add_rows_qty):
        for c_ in range(25):
            wb[sheet_name].cell(row=len(_df) + startrow - r_ - 1, column=c_ + 1).border = thin_border
    # перебираем скачанный датафрейм с добавленными строками с конца
    # (последняя строка остается - символ конца файла)
    for ir in range(len(_df) - 1, -1, -1):
        # проверка состояния флага (True - записать, False - удалить)
        #        print(ir, _df.iloc[ir][26])
        if _df.iloc[ir][26]:

            for ic in range(0, len(_df.iloc[ir]) - 2):
                wb[sheet_name].cell(startrow + ir, startcol + ic).value = _df.iloc[ir][ic]
                if ic in [4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 22]:
                    wb[sheet_name].cell(startrow + ir, startcol + ic).alignment = Alignment(horizontal='center')
                elif ic in [15, 16]:
                    wb[sheet_name].cell(startrow + ir, startcol + ic).alignment = Alignment(horizontal='center')
                    wb[sheet_name].cell(startrow + ir, startcol + ic).number_format = "# ##0"
                elif ic in [18, 19]:
                    wb[sheet_name].cell(startrow + ir, startcol + ic).alignment = Alignment(horizontal='center')
                    wb[sheet_name].cell(startrow + ir, startcol + ic).number_format = 'DD.MM.YYYY'
                elif ic in [23]:
                    wb[sheet_name].cell(startrow + ir, startcol + ic).alignment = Alignment(horizontal='center')
                    wb[sheet_name].cell(startrow + ir, startcol + ic).font = Font(color='808080')
                    wb[sheet_name].cell(startrow + ir, startcol + ic).fill = PatternFill('solid', fgColor="D9D9D9")
        else:
            if wb[sheet_name].cell(startrow + ir, 1).fill.start_color.index != '00000000':
                for i, row in enumerate(wb[sheet_name].iter_rows(min_row=startrow + ir, max_row=startrow + ir), 1):
                    for cell in row:
                        copy_cell(wb[sheet_name], cell.row, cell.column,
                                  wb['Избранное'], new_row_fav + cnt_fav, cell.column)
                cnt_fav += 1

            wb[sheet_name].delete_rows(startrow + ir, 1)
    if cnt_fav > 0:
        wb['Избранное'].conditional_formatting = ConditionalFormattingList()
    wb[sheet_name].cell(1, 1).value = today
    # форматирование цвета цены
    color_scale_rule = ColorScaleRule(start_type="min",
                                      start_color="57FFA3",  # зеленый
                                      end_type="max",
                                      end_color="FFEF9C")  # желтый

    # добавим этот градиент снова к столбцу `Q`.
    if len(_df) > 0:
        wb[sheet_name].conditional_formatting = ConditionalFormattingList()
        cells_range = 'R2:R' + str(_df.flag.sum() + 1)
        if (_df.flag.sum() + 1) >2:
            wb[sheet_name].conditional_formatting.add(cells_range, color_scale_rule)
        # форматирование цвета дат
            date_cells_range = 'T2:U' + str(_df.flag.sum() + 1)
        # redFill = PatternFill(fill_type='solid', start_color='FECECE', end_color='FECECE')
        # greenFill = PatternFill(fill_type='solid', start_color='D8E4BC', end_color='D8E4BC')
        # blueFill = PatternFill(fill_type='solid', start_color='B7DEE8', end_color='B7DEE8')
            redFill = PatternFill(fill_type='solid', bgColor='FECECE')
            greenFill = PatternFill(fill_type='solid', bgColor='D8E4BC')
            blueFill = PatternFill(fill_type='solid', bgColor='B7DEE8')
            diff_style_blue = DifferentialStyle(fill=blueFill)
            rule_blue = Rule(type="expression", dxf=diff_style_blue)
            rule_blue.formula = ["(T2-$A$1)<=0"]
            diff_style_green = DifferentialStyle(fill=greenFill)
            rule_green = Rule(type="expression", dxf=diff_style_green)
            rule_green.formula = ["(T2-$A$1)>7"]
            diff_style_red = DifferentialStyle(fill=redFill)
            rule_red = Rule(type="expression", dxf=diff_style_red)
            # rule_red.formula = ["=AND((T2-$A$1)<=7,(T2-$A$1)>0)"]
            rule_red.formula = ["(T2-$A$1)<=7"]

            wb[sheet_name].conditional_formatting.add(date_cells_range, rule_blue)
            wb[sheet_name].conditional_formatting.add(date_cells_range, rule_green)
            wb[sheet_name].conditional_formatting.add(date_cells_range, rule_red)

            cells_range = 'N2:N' + str(_df.flag.sum() + 1)
            pred_font = Font(color='0000FF')
            diff_style = DifferentialStyle(font=pred_font)
            rule = Rule(type="expression", dxf=diff_style)
            rule.formula = ["O2=1"]
            wb[sheet_name].conditional_formatting.add(cells_range, rule)
    wb.save(file_name)


def historical_processing(driver, in_process_df):

    # считываем данные о прошедших торгах из файла
    final_df = pd.read_excel(LOT_FILENAME, sheet_name='Прошедшие', engine='openpyxl',
                             usecols=[2])

    df_for_save = pd.DataFrame(columns = ['lot_tag', 'addr_norm', 'adm_area', 'mun_district', 'metro_name_gkh', 'metro_km',
                           'metro_minutes', 'addr_floor', 'total_floors', 'year_exp_w', 'реновация', 'ремонт',
                           'qty_rooms', 'is_qty_rooms_predicted', 'obj_square',
                           'start_price', 'start_price_m2', 'final_price', 'final_price_m2', 'delta_price', 'status',
                           'auct_date', 'investMSK_URL', 'lot_URL', 'platform_name',
                           'timestamp', 'building_page'])

    today = dt.datetime.today()
    in_process_df.auct_date = pd.to_datetime(in_process_df.auct_date)
    url_lst = in_process_df['investMSK_URL'].to_list()
    unreadable_urls = []
    print('Обработка данных со статусом InProcess ')
    logging.info('Обработка данных со статусом InProcess ')
    for url in tqdm(url_lst, position=0):
        # считываем страницу с заданным URL, при нескачивании в течении 10 секунд добавляем ссылку в список нескачанных
        try:

            lot_status_torgi = 0
            driver, lot_status_data = control_lot(driver, url)
            if (lot_status_data['status'] == 'ПРИЕМ ЗАЯВОК ЗАВЕРШЕН') \
                    and (in_process_df.loc[in_process_df.investMSK_URL == url].iloc[0]['platform_name'] == 'torgi.gov.ru'):
                if in_process_df.loc[in_process_df.investMSK_URL == url].iloc[0]['auct_date'] <= today:

                    driver, lot_status_torgi, final_price = \
                        is_lot_finished(driver, in_process_df.loc[in_process_df.investMSK_URL == url].iloc[0]['lot_URL'])

            if lot_status_torgi == -1:
                lot_status_data['status'] = 'ПРИЗНАНЫ НЕСОСТОЯВШИМИСЯ'
            elif lot_status_torgi == 1:
                lot_status_data['status'] = 'ПРИЗНАНЫ СОСТОЯВШИМИСЯ'
                lot_status_data['final_price'] = final_price
            if (lot_status_data['status'] == 'ПРИЗНАНЫ НЕСОСТОЯВШИМИСЯ') or \
                    (lot_status_data['status'] == 'ПРИЗНАНЫ СОСТОЯВШИМИСЯ') or \
                    (lot_status_data['status'] == 'ОТМЕНЕНЫ'):
            # лоты с информацией о проведении добавляем в df для записи в прошедшие
            # и помечаем флагом False для уадления из инпроцесс
                in_process_df.loc[in_process_df.investMSK_URL == url, 'flag'] = False
                row_to_add = in_process_df.loc[in_process_df.investMSK_URL == url].iloc[0]
                row_to_add = pd.concat([row_to_add, pd.Series(lot_status_data)])
              #  row_to_add = row_to_add.drop(columns = ['limit', 'deadline', 'flag', 'flag_duplication'])
                try:
                    row_to_add['final_price_m2'] = round(row_to_add['final_price'] / row_to_add['obj_square'], 2)
                    row_to_add['delta_price'] = round(
                        (row_to_add['final_price'] - row_to_add['start_price']) / row_to_add['start_price'], 4)
                except (Exception,):
                    pass
                df_for_save = pd.concat([df_for_save, row_to_add.to_frame().T], ignore_index=True)
            else:
                # if lot_status_data['status']:
                #     print(lot_status_data['status'])
                in_process_df.loc[in_process_df.investMSK_URL == url, 'flag'] = True

        except Exception as e:
            print(f'Нечитаемая ссылка {url}')
            logging.info(f'Нечитаемая ссылка {url}')
            unreadable_urls.append(url)
            print('Ошибка:\n', traceback.format_exc())
            logging.info(f'Ошибка:\n {traceback.format_exc()}')
            continue

    print('Пропущенных объектов: ', len(unreadable_urls))
    logging.info(f'Пропущенных объектов: {len(unreadable_urls)}')
    df_for_save = df_for_save[['lot_tag', 'addr_norm', 'adm_area', 'mun_district', 'metro_name_gkh', 'metro_km',
                 'metro_minutes', 'addr_floor', 'total_floors', 'year_exp_w', 'реновация', 'ремонт',
                 'qty_rooms', 'is_qty_rooms_predicted', 'obj_square',
                 'start_price', 'start_price_m2', 'final_price', 'final_price_m2', 'delta_price', 'status',
                 'auct_date', 'investMSK_URL', 'lot_URL', 'platform_name',
                 'timestamp', 'building_page']]
    print('Запись данных в файл MosInvestData_v2.xlsx')
    logging.info('Запись данных в файл MosInvestData_v2.xlsx')
    try:
        update_spreadsheet_hist(LOT_FILENAME, df_for_save, startcol=2, startrow=len(final_df)+2,
                            sheet_name='Прошедшие')
    except:
        print('Закройте файл MosInvestData_v2.xlsx и нажмите Enter')
        logging.info('Закройте файл MosInvestData_v2.xlsx и нажмите Enter')
        input()
        update_spreadsheet_hist(LOT_FILENAME, df_for_save, startcol=2, startrow=len(final_df) + 2,
                        sheet_name='Прошедшие')
    print('Обновление данных о прошедших торгах завершено.')
    logging.info('Обновление данных о прошедших торгах завершено.')
    return in_process_df


def actual_moskowinvest():
    # global gkh_df
    # global date_to
    # global clf
    # global MID_df
    start_time = time.time()
    print('Обновление данных об актуальных лотах ')
    # считываем данные о прошедших торгах из файла
    print('Попытка получить данные с сайта InvestMoscow')
    # запускаем Chrome браузер в оконном режиме для получения информации с сайта
    # в headless режиме - крашится
    driver = start_browser_for_parse()
    driver.implicitly_wait(30)
    # url до 26.04.2023 'https://investmoscow.ru/tenders?pageNumber=1&pageSize=100&orderBy=CreateDate&orderAsc=false&objectTypes=7&tenderTypes=13&tenderStatuses=1&tradeForms=45001'
    main_page_url = 'https://investmoscow.ru/tenders?pageNumber=1&pageSize=10&orderBy=RequestEndDate&orderAsc=true&objectTypes=nsi:41:30011568&objectKinds=nsi:tender_type_portal:13&tenderStatus=nsi:tender_status_tender_filter:1&timeToPublicTransportStop.noMatter=true'
    driver.get(main_page_url)
    wait = WebDriverWait(driver, 30)
    try:
        wait.until(
            EC.visibility_of_element_located((By.XPATH,
                                              '//span[@class="uid-text-accent"]')))
        main_page_soup = BeautifulSoup(driver.page_source, 'html.parser')
        qty_lots_str = main_page_soup.find('span', class_="uid-text-accent").text.strip()
    except (Exception, ):
        print('Сайт не отвечает. Попробуйте позже.')
        logging.info('Сайт не отвечает. Попробуйте позже.')
        driver.quit()
        return
    qty_lots = int(re.sub("\D", "", qty_lots_str))
    print('Всего объектов: ', qty_lots)
    logging.info(f'Всего объектов: {qty_lots}')
    if qty_lots == 0:
        return
    url_lst = []
    # считываем данные страниц по актуальным лотам торгам
    for i in trange(int(qty_lots / 100) + 1):
        if len(url_lst) >= qty_lots:
            break
        cur_page_url = 'https://investmoscow.ru/tenders?pageNumber=' + str(i + 1) + '&pageSize=100&' + \
                       'orderBy=RequestEndDate&orderAsc=true&objectTypes=nsi:41:30011568&' + \
                       'objectKinds=nsi:tender_type_portal:13&tenderStatus=nsi:tender_status_tender_filter:1' + \
                        '&timeToPublicTransportStop.noMatter=true'
        driver, url_page_lst = get_url_list_from_page(driver, cur_page_url)
        url_lst = url_lst + url_page_lst
        #url_lst .append(url_page_lst)
    # создаем датафрейм с данными лотов
    lot_df = pd.DataFrame(columns=LOT_FIELDS)
    print("time elapsed: {:.2f}s".format(time.time() - start_time))
    unreadable_urls = []
    error_url_cnt = 0
    driver.implicitly_wait(10)
#---> тут можно добавить ограничение по парсингу первой части
    for url in tqdm(url_lst):
        # считываем страницу с заданным URL, при нескачивании в течении 10 секунд добавляем ссылку в список нескачанных
        try:
            # page = requests.get(url, timeout=10)
            # soup = BeautifulSoup(page.text, "html.parser")
            driver, lot_single = parse_lot(driver, url)
            #lot_single = parse_lot(soup, is_actual=True)
            lot_single['investMSK_URL'] = url
            if not (pd.isna(lot_single['start_price'])):
                lot_df = pd.concat([lot_df, lot_single.to_frame().T], ignore_index = True)
            else:
                print(f'Нечитаемая ссылка N{error_url_cnt} {url}')
                print('lot_tag ' + lot_single['lot_tag'])
                logging.info(f'Нечитаемая ссылка N{error_url_cnt} {url}')
                logging.info(f"lot_tag  {lot_single['lot_tag']}")
                unreadable_urls.append(url)
                # print('Ошибка: ', traceback.format_exc())
                # logging.info(f'Ошибка: {traceback.format_exc()}')
                error_url_cnt += 1
                continue
        except Exception as e:
            print(f'Нечитаемая ссылка N{error_url_cnt} {url}')
            print('lot_tag ' + lot_single['lot_tag'])
            logging.info(f'Нечитаемая ссылка N{error_url_cnt} {url}')
            logging.info(f"lot_tag  {lot_single['lot_tag']}")
            unreadable_urls.append(url)
            # print('Ошибка: ', traceback.format_exc())
            # logging.info(f'Ошибка: {traceback.format_exc()}')
            error_url_cnt += 1
            continue
    print('Всего новых объектов: ', qty_lots)
    print('Скачалось объектов: ', len(lot_df))
    print('Пропущенных объектов: ', len(unreadable_urls))
    logging.info('Всего новых объектов: ' + str(qty_lots))
    logging.info('Скачалось объектов: ' + str(len(lot_df)))
    logging.info('Пропущенных объектов: ' + str(len(unreadable_urls)))
    # повторно пытаемся скачать ранее нескачанные ссылки с таймаутом 60 секунд
    if len(unreadable_urls) > 0:
        lot_df = reparsing_unreadable_urls(driver, lot_df, qty_lots, unreadable_urls)

    # исключаем строки, в прием заявок по которым окончен ранее или сегодня
    today = pd.to_datetime('today')
    lot_df = lot_df[lot_df['deadline'] > today]

    # очистка датафрейма от дублирующихся лотов
    if len(lot_df) > 0:
        lot_df = drop_duplicated_lots(lot_df)
    else:
        print('Новые данные отсутвуют')
        logging.info('Новые данные отсутвуют')
    # нормируем и распознаем адреса (если есть новые)
    if len(lot_df) > 0:
        norm_df = recognize_and_normalize_addresses(lot_df)
        # после распознавания адресов - уточняем данные о метро и этажности
        norm_df = is_metro_and_floor_data_complete(norm_df)
    else:
        norm_df = pd.DataFrame()
    # в случае, если в карточке отсутствует указание на этажность дома - заполняем по данным ЖКХ
    if len(norm_df) > 0:
        len_norm_df = 1
        norm_df = fill_spaces_in_data(norm_df)

        df_for_save = norm_df[['lot_tag', 'addr_norm', 'adm_area', 'mun_district', 'metro_name_gkh', 'metro_km',
                           'metro_minutes', 'addr_floor', 'total_floors', 'year_exp_w', 'реновация', 'ремонт',
                           'qty_rooms', 'obj_square',
                           'start_price', 'start_price_m2', 'deadline',
                           'auct_date', 'investMSK_URL', 'lot_URL', 'platform_name',
                           'timestamp', 'building_page']]
        df_for_save = df_for_save.rename(columns={'metro_name_gkh': 'metro_name'})

        # обновляем датафрейм MID.xls
        MID_df = pd.read_excel(LOT_FILENAME, sheet_name='Активные', engine='openpyxl',
                               converters={"lot_tag": str}
                               )

        MID_df = MID_df.dropna(subset=['lot_tag'])
    #    MID_df = MID_df[:-1]
        df_for_save['Примечания'] = np.nan

        df_for_save['flag'] = np.nan
        MID_df['flag'] = np.nan
        # в df_for_save flag = -1 для строк, отсутсвующих в MID, -2 для строк, в которых изменен дедлайн или дата торгов
        df_for_save['flag'] = df_for_save.apply(lambda x: compare_with_existing(x), axis=1)
        # в датафрейм MID добавляем новые строки (flag = -1)
        add_cells_qty = df_for_save[df_for_save.flag == -1].shape[0] # количество новых строк для эксельки
        add_cells_qty += df_for_save[df_for_save.flag == -2].shape[0]
        MID_df = pd.concat([MID_df, df_for_save[df_for_save.flag == -1]], ignore_index=True)
        MID_df = pd.concat([MID_df, df_for_save[df_for_save.flag == -2]], ignore_index=True)
        # дополняем колонкой, обозначающей дубликаты // инверсный флаг - те строки, которые надо оставить помечаются True
        MID_df['flag_duplication'] = ~MID_df.duplicated(subset=['lot_tag'], keep='last')
        # устанавливаем флаг актуальности (если дата подачи сегодня или  ранее - запись не актуальна)
        MID_df.flag = MID_df.deadline > today
        # выбираем лоты с неактуальными датами для добавления в InProcess
        in_process_add_df = MID_df[~MID_df.flag].copy()
        # объединяем флаги для последующего удаления строк в Эксельке - любой False удалится
        MID_df.flag = MID_df.flag & MID_df.flag_duplication
        MID_df['lot_tag'] = MID_df['lot_tag'].apply(lambda x: str(x).strip())

        # ---------> заполняем пропуски в количестве комнат
        MID_df = fill_qty_rooms_with_predictions(MID_df)
        MID_df['metro_minutes'] = pd.to_numeric(MID_df['metro_minutes'], errors='coerce')
        # подготовка данных MID для объеднинения с базой ЦИАН
        cut_labels_4 = ['<10 мин', '<20 мин', '<30 мин', '>30 мин']
        cut_bins = [0, 10, 20, 30, 3000]
        MID_df['metro_min'] = pd.cut(MID_df['metro_minutes'],
                                     bins=cut_bins,
                                     labels=cut_labels_4)
        MID_df['is_renov'] = MID_df['реновация'].apply(lambda x: 'Нет' if (x == 'новый дом') and (x == 'нет в плане')
                                                                          and (x == 'N/A') else 'Да')
        rooms_labels_4 = ['1', '2', '3', '4+']
        rooms_bins = [0, 1, 2, 3, 20]
        MID_df['rooms_qty'] = pd.cut(MID_df['qty_rooms'],
                                     bins=rooms_bins,
                                     labels=rooms_labels_4)
        MID_df['metro_name_lower'] = MID_df.metro_name.apply(metro_prep)

        # формирование окончательного датафрейма
        MID_df = MID_df[['lot_tag', 'addr_norm', 'adm_area', 'mun_district', 'metro_name', 'metro_km', 'metro_minutes',
                         'addr_floor', 'total_floors', 'year_exp_w', 'реновация', 'ремонт', 'qty_rooms',
                         'is_qty_rooms_predicted',
                         'obj_square', 'start_price', 'start_price_m2', 'limit', 'deadline', 'auct_date',
                         'investMSK_URL', 'lot_URL',
                         'platform_name', 'timestamp', 'building_page', 'Примечания', 'flag']]

        print('Запись данных по активным лотам в файл MosInvestData_v2.xlsx')
        logging.info('Запись данных по активным лотам в файл MosInvestData_v2.xlsx')
        try:
            update_spreadsheet_w_drop(LOT_FILENAME,
                                  MID_df, startcol=2, startrow=2, add_rows_qty=add_cells_qty, sheet_name='Активные')
        except:
            print('Закройте файл MosInvestData_v2.xlsx и нажмите Enter')
            logging.info('Закройте файл MosInvestData_v2.xlsx и нажмите Enter')
            input()
            update_spreadsheet_w_drop(LOT_FILENAME,
                                      MID_df, startcol=2, startrow=2, add_rows_qty=add_cells_qty, sheet_name='Активные')

        print('Обработка новых лотов')
        logging.info('Обработка новых лотов')
        # ------------> обработка листа с новыми объектами
        new_MID_df = pd.read_excel(LOT_FILENAME, sheet_name='NEW!!!', engine='openpyxl',
                                   converters={"lot_tag": str, "metro_minutes" : str}
                                   )
        new_MID_df = new_MID_df.dropna(subset=['lot_tag'])
    #    new_MID_df = new_MID_df[:-1]
        # существующим записям присваиваем флаг False для удаления
        new_MID_df['flag'] = False
        # добавляем новые и измененные данные
        new_MID_df = pd.concat([new_MID_df, df_for_save[df_for_save.flag == -1]], ignore_index=True)
        new_MID_df = pd.concat([new_MID_df, df_for_save[df_for_save.flag == -2]], ignore_index=True)
        if len(new_MID_df) > 0:
            add_cells_qty = df_for_save[df_for_save.flag == -1].shape[0] + df_for_save[df_for_save.flag == -2].shape[0]
            new_MID_df.flag = new_MID_df.flag < 0
            new_MID_df['metro_minutes'] = pd.to_numeric(new_MID_df['metro_minutes'], errors='coerce')

            # подготовка данных MID для объеднинения с базой ЦИАН
            cut_labels_4 = ['<10 мин', '<20 мин', '<30 мин', '>30 мин']
            cut_bins = [0, 10, 20, 30, 3000]
            new_MID_df['metro_min'] = pd.cut(new_MID_df['metro_minutes'],
                                         bins=cut_bins,
                                         labels=cut_labels_4)
            new_MID_df['is_renov'] = new_MID_df['реновация'].apply(lambda x: 'Нет' if (x == 'новый дом') and (x == 'нет в плане')
                                                                              and (x == 'N/A') else 'Да')
            new_MID_df = fill_qty_rooms_with_predictions(new_MID_df)
            rooms_labels_4 = ['1', '2', '3', '4+']
            rooms_bins = [0, 1, 2, 3, 20]
            new_MID_df['rooms_qty'] = pd.cut(new_MID_df['qty_rooms'],
                                         bins=rooms_bins,
                                         labels=rooms_labels_4)
            new_MID_df['metro_name_lower'] = new_MID_df.metro_name.apply(metro_prep)
            new_MID_df = new_MID_df[['lot_tag', 'addr_norm', 'adm_area', 'mun_district', 'metro_name', 'metro_km',
                                     'metro_minutes', 'addr_floor', 'total_floors', 'year_exp_w', 'реновация', 'ремонт',
                                     'qty_rooms', 'is_qty_rooms_predicted',
                                     'obj_square', 'start_price', 'start_price_m2', 'limit', 'deadline', 'auct_date',
                                     'investMSK_URL', 'lot_URL',
                                     'platform_name', 'timestamp', 'building_page', 'Примечания', 'flag']]
            print('Запись данных по новым лотам в файл MosInvestData_v2.xlsx')
            logging.info('Запись данных по новым лотам в файл MosInvestData_v2.xlsx')
            try:
                update_spreadsheet_w_drop(LOT_FILENAME, new_MID_df, sheet_name='NEW!!!',
                                          startcol=2, startrow=2, add_rows_qty=add_cells_qty)
            except (Exception, ):
                print('Закройте файл MosInvestData_v2.xlsx и нажмите Enter')
                logging.info('Закройте файл MosInvestData_v2.xlsx и нажмите Enter')
                input()
                update_spreadsheet_w_drop(LOT_FILENAME, new_MID_df, sheet_name='NEW!!!',
                                          startcol=2, startrow=2, add_rows_qty=add_cells_qty)
        print('Обработка лотов, по которым закончен прием предложений')
        logging.info('Обработка лотов, по которым закончен прием предложений')
    else:
        print('Новые объекты отсутсвуют')
        logging.info('Новые объекты отсутсвуют')
        len_norm_df = 0
    # ------------> обработка листа с объектами InProcess
    in_process_df = pd.read_excel(LOT_FILENAME, sheet_name='InProcess', engine='openpyxl',
                                  converters={"lot_tag": str, "metro_minutes": str})
    in_process_df = in_process_df.dropna(subset=['lot_tag'])
    start_len = len(in_process_df)
    # добавляем новые и измененные данные
    if len_norm_df == 1:
        in_process_df = pd.concat([in_process_df, in_process_add_df], ignore_index=True)
    in_process_df['flag'] = np.nan
    if len(in_process_df) > 0:
        in_process_df = historical_processing(driver, in_process_df)
        print('Сохраняются данные лотов, по которым закончен прием предложений')
        logging.info('Сохраняются данные лотов, по которым закончен прием предложений')
        add_cells_qty = len(in_process_df) - start_len
        in_process_df = in_process_df[
            ['lot_tag', 'addr_norm', 'adm_area', 'mun_district', 'metro_name', 'metro_km', 'metro_minutes',
             'addr_floor', 'total_floors', 'year_exp_w', 'реновация', 'ремонт', 'qty_rooms',
             'is_qty_rooms_predicted',
             'obj_square', 'start_price', 'start_price_m2', 'limit', 'deadline', 'auct_date',
             'investMSK_URL', 'lot_URL',
             'platform_name', 'timestamp', 'building_page', 'Примечания', 'flag']]
        try:
            update_spreadsheet_w_drop(LOT_FILENAME, in_process_df, sheet_name='InProcess',
                                      startcol=2, startrow=2, add_rows_qty=add_cells_qty)
        except (Exception, ):
            print('Закройте файл MosInvestData_v2.xlsx и нажмите Enter')
            logging.info('Закройте файл MosInvestData_v2.xlsx и нажмите Enter')
            input()
            update_spreadsheet_w_drop(LOT_FILENAME, in_process_df, sheet_name='InProcess',
                                      startcol=2, startrow=2, add_rows_qty=add_cells_qty)
    driver.quit()
    try:
        driver.quit()
    except (Exception,):
        pass
    print('Обновление данных завершено.')
    logging.info('Обновление данных завершено.')


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    #
    # test_url_list=['https://torgi.gov.ru/new/public/notices/view/21000005000000010042']
    # driver = start_browser_for_parse()
    # for url_ in test_url_list:
    #     driver, lot_status, final_price = is_lot_finished(driver, url_)
    #     print(lot_status, final_price)
    LOT_FIELDS = ['lot_tag', 'auct_type', 'object_type', 'cadastr_num', 'addr_string',
                  'flat_num', 'qty_rooms', 'addr_floor', 'total_floors', 'obj_square', 'start_price',
                  'start_price_m2','deposit', 'auct_step', 'auct_form', 'start_applications_date',
                  'finish_application_date', 'participant_selection_date',
                  'bidding_date', 'results_date', 'roseltorg_url', 'torgi_URL',
                  'metro_station', 'inf_sales', 'inf_food_service', 'inf_education',
                  'inf_cult_and_sport', 'inf_consumer_services', 'documentation_expl',
                  'documentation_photo', 'investmoscow_url']
    LOT_FILENAME = r'..\realty_model_lot_data.xlsx'
    SCROLL_PAUSE_TIME = 0.5
    current_date = datetime.datetime.now()
    current_date_string = current_date.strftime('%y_%m_%d_%H_%M')
    print(current_date_string)
    logging.basicConfig(
        level=logging.INFO,
        filename="model_log_" + current_date_string + ".log",
        format="%(asctime)s - %(module)s - %(levelname)s - %(funcName)s: %(lineno)d - %(message)s",
        datefmt='%H:%M:%S',
    )

    logging.info('Hello')
    clf = tree.DecisionTreeClassifier(criterion='entropy', max_depth=10)
    MID_df = pd.DataFrame()
    gkh_df = pd.DataFrame()
    date_to = ''
    morph_vocab = MorphVocab()
    addr_extractor = AddrExtractor(morph_vocab)
    choice = 'q'
    min_dict = dict({0: ['&foot_min=10', '<10 мин', '&only_foot=2'],
                     1: ['&foot_min=20', '<20 мин', '&only_foot=2'],
                     2: ['&foot_min=30', '<30 мин', '&only_foot=2'],
                     3: ['', '>30 мин', '']
                     })
    room_dict = dict({0: ['&room1=1', '1'],
                      1: ['&room2=1', '2'],
                      2: ['&room3=1', '3'],
                      3: ['&room4=1&room5=1&room6=1', '4+']
                      })
    try:
        actual_moskowinvest()
    except Exception as e:
        print('Ошибка: ', traceback.format_exc())
        logging.info(f'Ошибка: {traceback.format_exc()}')

