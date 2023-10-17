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

# -
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
    """Функция переключения по вкладкам внутри сраницы объекта на investmoscow.ru
    :param tab_name: надпись на вкладке
    :return driver после открытия нужной вкладки"""
    element = driver.find_element(By.XPATH, "//div[text()='"+tab_name+"']")
    body = driver.find_element_by_css_selector('body')

    try:
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//div[text()='"+tab_name+"']")))
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
                EC.element_to_be_clickable((By.XPATH, "//div[text()='"+tab_name+"']")))
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
                        EC.element_to_be_clickable((By.XPATH, "//div[text()='"+tab_name+"']")))
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
    obj_dict['investmoscow_url'] = url

    # выбираем данные по площади и местоположению
    try:
        obj_dict['obj_square'] = float(
            str.replace(full_info.find('span', class_="uid-text-accent uid-mr-16").text, ',', '.').split(' ')[0])
    except Exception as exc:
        print('Ошибка определения площади объекта ', url, exc)
        logging.info(f"Лот {url} - Ошибка определения площади объекта")

    labels_info = full_info.find_all('div', class_="tender-card__header-label")
    for label in labels_info:
        if label.text.strip().startswith('#'):
            try:
                obj_dict['lot_tag'] = label.text.strip()
            except Exception as exc:
                obj_dict['lot_tag'] = np.nan
                print('Ошибка определения номера лота объекта ', url, exc)
                logging.info(f"Лот {url} - Ошибка определения номера лота объекта")
        elif label.text.strip().startswith('Тип объекта:'):
            try:
                obj_dict['object_type'] = label.text.split(':')[-1].strip()
            except Exception as exc:
                obj_dict['object_type'] = np.nan
                print('Ошибка определения типа объекта ', url, exc)
                logging.info(f"Лот {url} - Ошибка определения типа объекта")
        elif label.text.strip().startswith('Вид торгов:'):
            try:
                obj_dict['auct_type'] = label.text.split(':')[-1].strip()
            except Exception as exc:
                obj_dict['auct_type'] = np.nan
                print('Ошибка определения вида торгов ', url, exc)
                logging.info(f"Лот {url} - Ошибка определения вида торгов объекта")

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
                obj_dict['addr_floor'] = int(block.find('div', class_="subject-table-text").text
                                             .strip('№').strip('N').strip())
            except Exception as exc:
                if block.find('div', class_="subject-table-text").text != 'Не указано':
                    print(f"Лот {url} - Ошибка определения этажа объекта")
                    logging.info(f"Лот {url} - Ошибка определения этажа объекта")

        elif label_text == 'Этажность дома:':
            try:
                obj_dict['total_floors'] = int(block.find('div', class_="subject-table-text").text)
            except Exception as exc:
                if block.find('div', class_="subject-table-text").text != 'Не указано':
                    print(f"Лот {url} - Ошибка определения этажности объекта")
                    logging.info(f"Лот {url} - Ошибка определения этажности объекта")

        elif label_text == 'Тип объекта:':
            if not obj_dict['object_type']:
                obj_dict['object_type'] = block.find('div', class_="subject-table-text").text

        elif label_text == 'Кадастровый номер:':
            obj_dict['cadastr_num'] = block.find('div', class_="subject-table-text").text

        elif label_text == 'Номер квартиры:':
            try:
                obj_dict['flat_num'] = int(block.find('div', class_="subject-table-text").text
                                           .strip('№').strip('N').strip())
            except Exception as exc:
                if block.find('div', class_="subject-table-text").text != 'Не указано':
                    print(f"Лот {url} - Ошибка определения номера квартиры объекта")
                    logging.info(f"Лот {url} - Ошибка определения номера квартиры объекта")

        elif label_text == 'Количество комнат:':
            try:
                obj_dict['qty_rooms'] = int(block.find('div', class_="subject-table-text").text)
            except Exception as exc:
                if block.find('div', class_="subject-table-text").text != 'Не указано':
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

        elif label_text == 'Дата начала приёма заявок:':
            obj_dict['start_applications_date'] = pd.to_datetime(
                block.find('div', class_="subject-table-text").text.split()[0],
                infer_datetime_format=True, format='%d.%m.%Y', errors='coerce')

        elif label_text == 'Дата окончания приёма заявок:':
            obj_dict['finish_application_date'] = pd.to_datetime(
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
            try:
                obj_dict['roseltorg_url'] = block.find('a')['href']
            except Exception as exc:
                pass

        elif label_text == 'Ссылка на torgi.gov.ru:':
            try:
                obj_dict['torgi_url'] = block.find('a')['href']
            except Exception as exc:
                pass


    # переключаемся на "Дополнительная информация"
    driver = go_to_procedure_detail(driver, 'Дополнительная информация')

    wait.until(
        EC.visibility_of_element_located((By.XPATH,
                                          '//div[@class="tender__content"]')))
    full_info = BeautifulSoup(
        driver.find_element(By.XPATH, '//div[@class="extra-info"]').get_attribute('innerHTML'),
        "html.parser")
    #ищем блок "Транспортная доступность", проверяем наличие картинки метро, сохраняем название станции
    trans_desc = full_info.find('div', class_="extra-info-transport")
    metro_logo = trans_desc.find('img')['src']
    if re.search('metro', metro_logo):
        obj_dict['metro_station'] = trans_desc.div.div.div.text

    # ищем блок "Инфраструктура"
    infra_desc = full_info.find('div', class_="extra-info-block")
    t1 = infra_desc.find_all('a', class_="extra-info-block-item uid-portal-link")
    for block in t1:
        label_text = block.find('span').text.strip()
        infra_item_meaning = int(block.find('div', class_="extra-info-block-item-text").text.strip())
        if label_text == "Торговля":
            obj_dict['inf_sales'] = infra_item_meaning
        elif label_text == "Общественное питание":
            obj_dict['inf_food_service'] = infra_item_meaning
        elif label_text == "Образование":
            obj_dict['inf_education'] = infra_item_meaning
        elif label_text == "Культура и спорт":
            obj_dict['inf_cult_and_sport'] = infra_item_meaning
        elif label_text == "Бытовое обслуживание":
            obj_dict['inf_consumer_services'] = infra_item_meaning
        elif label_text == "Здравоохранение":
            obj_dict['inf_health_care'] = infra_item_meaning
        else:
            print(f"Лот {url} - неучтенная инфраструктура - ", label_text, infra_item_meaning)
            logging.info(f"Лот {url} - неучтенная инфрастурктура - "
                         + label_text + str(infra_item_meaning))

    return driver, pd.Series(obj_dict)

# -
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
    adr_apart_num = np.nan
    for _, paramValue in enumerate(addr_extractor.find(line).fact.parts):
        p0 = addr_extractor.find(line).fact.parts
    try:
        for p in p0:
            if p.type == 'город':
                if p.value != 'Москва' and p.value != 'Москвы':
                    adr_street += str(', г. ' + p.value)
            elif p.type == 'деревня':
                adr_street += str(', д. ' + p.value)
            elif p.type == 'микрорайон':
                adr_street += str(', мкр. ' + p.value)
            elif p.type == 'бульвар':
                adr_street += str(', б-р. ' + p.value)
            elif p.type == 'дачный поселок':
                adr_street += str(', дп. ' + p.value)
            elif p.type == 'шоссе':
                adr_street += str(', ш. ' + p.value)
            elif p.type == 'улица':
                adr_street += str(', ул. ' + p.value)
            elif p.type == 'переулок':
                adr_street += str(', пер. ' + p.value)
            elif p.type == 'квартал':
                adr_street += str(', кв-л. ' + p.value)
            elif p.type == 'проезд':
                adr_street += str(', проезд. ' + p.value)
            elif p.type == 'проспект':
                adr_street += str(', пр-кт. ' + p.value)
            elif p.type == 'аллея':
                adr_street += str(', аллея. ' + p.value)
            elif p.type == 'площадь':
                adr_street += str(', пл. ' + p.value)
            elif p.type == 'набережная':
                adr_street += str(', наб. ' + p.value)
            elif p.type == 'село':
                adr_street += str(', с. ' + p.value)
            elif p.type == 'дом':
                adr_house = str(', д. ' + p.value)
            elif p.type == 'корпус':
                adr_house = adr_house + str(', к. ' + p.value)
            elif p.type == 'строение':
                adr_house = adr_house + str(', стр. ' + p.value)
            elif p.type == 'этаж':
                adr_floor = p.value
            elif p.type == 'квартира':
                adr_apart_num = p.value
        adr_street = 'г. Москва' + adr_street
        ret = [adr_street, adr_house, adr_floor, adr_apart_num]
    except:
        ret = ['не распознан', 'не распознан', 'не распознан', 'не распознан']
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
    adr_apart_num = np.nan
        # загружаем свой словарь обработки нестандартных адресов
    # TODO: заменить на csv
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
            elif p.type == 'деревня':
                adr_street += str(', д. ' + p.value)
            elif p.type == 'микрорайон':
                adr_street += str(', мкр. ' + p.value)
            elif p.type == 'бульвар':
                adr_street += str(', б-р. ' + p.value)
            elif p.type == 'дачный поселок':
                adr_street += str(', дп. ' + p.value)
            elif p.type == 'поселок':
                adr_street += str(', п. ' + p.value)
            elif p.type == 'шоссе':
                adr_street += str(', ш. ' + p.value)
            elif p.type == 'квартал':
                adr_street += str(', кв-л. ' + p.value)
            elif p.type == 'улица':
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

            elif p.type == 'переулок':
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

            elif p.type == 'проезд':
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

            elif p.type == 'проспект':
                adr_street += str(', пр-кт. ' + p.value)
            elif p.type == 'аллея':
                adr_street += str(', аллея. ' + p.value)
            elif p.type == 'площадь':
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
            elif p.type == 'набережная':
                adr_street += str(', наб. ' + p.value)
            elif p.type == 'село':
                adr_street += str(', с. ' + p.value)
            elif p.type == 'дом':
                adr_house = str(', д. ' + p.value)
            elif p.type == 'корпус':
                adr_house = adr_house + str(', к. ' + p.value)
            elif p.type == 'строение':
                adr_house = adr_house + str(', стр. ' + p.value)
            elif p.type == 'этаж':
                adr_floor = p.value
            elif p.type == 'квартира':
                adr_apart_num = p.value

        adr_street = re.sub('ё', 'е', adr_street)
        adr_street = 'г. Москва' + adr_street
        ret = [adr_street, adr_house, adr_floor, adr_apart_num]
    except:
        return ['не распознан', 'не распознан', 'не распознан', 'не распознан']
    return ret

# -
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
        elif p.type == 'деревня':
            adr_street += str(', д. ' + p.value)
        elif p.type == 'микрорайон':
            adr_street += str(', мкр. ' + p.value)
        elif p.type == 'бульвар':
            adr_street += str(', б-р. ' + p.value)
        elif p.type == 'дачный поселок':
            adr_street += str(', дп. ' + p.value)
        elif p.type == 'поселок':
            adr_street += str(', п. ' + p.value)
        elif p.type == 'шоссе':
            adr_street += str(', ш. ' + p.value)
        elif p.type == 'квартал':
            adr_street += str(', кв-л. ' + p.value)
        elif p.type == 'улица':
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
        elif p.type == 'переулок':
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
        elif p.type == 'проезд':
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
        elif p.type == 'проспект':
            adr_street += str(', пр-кт. ' + p.value)
        elif p.type == 'аллея':
            adr_street += str(', аллея. ' + p.value)
        elif p.type == 'площадь':
            adr_street += str(', пл. ' + p.value)
        elif p.type == 'набережная':
            adr_street += str(', наб. ' + p.value)
        elif p.type == 'село':
            adr_street += str(', с. ' + p.value)
        elif p.type == 'дом':
            adr_house = str(', д. ' + p.value)
        elif p.type == 'корпус':
            adr_house = adr_house + str(', к. ' + p.value)
        elif p.type == 'строение':
            adr_house = adr_house + str(', стр. ' + p.value)
    adr_street = re.sub('ё', 'е', adr_street)
    adr_street = 'г. Москва' + adr_street
    ret = adr_street + adr_house
    return ret

# -
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

# -
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
    gkh_df.to_csv(GKH_BASE_FILENAME, index=False)
    clean_gkh_add(r'..\new_buildings.xlsx', sheet_name='Sheet1')


def update_gkh_base(address, res):
    """ обновление информации по адресу в базе ЖКХ.
        Не забыть скачать базу до и сохранить в файл после работы  """
    global gkh_df
    gkh_df = pd.concat([gkh_df, pd.DataFrame.from_records([res])] ,ignore_index=True)
    gkh_df.drop_duplicates(subset=['gkh_address'], keep='last', inplace=True, ignore_index=True)
    gkh_df.reset_index(drop=True, inplace=True)

# -
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

# -
def metro_and_floor_data(addr_norm):
    """ Функция поиска данных об адресе на ресурсе flatInfo.ru (этажность и расстояние до метро)
    Функция дополнительно обновляет базу ЖКХ update_gkh_base    """
    global gkh_df

    print(addr_norm)
    logging.info(str(addr_norm))
    #res = [np.nan for _ in range(5)]
    res = {GKH_FIELDS[i]: np.nan for i in range(len(GKH_FIELDS))}
    try:
        driver = start_browser_for_parse()
        driver.get('https://flatinfo.ru')
        element = driver.find_element(By.XPATH, "//div[@class='search-home input-group search-home_show']/input[1]")
        adr1 = re.sub('проезд.', 'проезд', addr_norm)
        adr1 = re.sub('пр-кт.', 'проспект', adr1)
        adr1 = re.sub('г. зеленоград, к.', 'г. зеленоград,', adr1)
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
        if re.search('h_info', cur_url) is None:
            print('адреса не совпадают')
            logging.info('адреса не совпадают')
            print('Ввести правильный url?')
            is_cont = input()
            if is_cont in ['Y', 'y']:
                print('Введите правильный url')
                new_url = input()
                driver.get(new_url)
                cur_url = new_url

        if re.search('h_info', cur_url) is not None:
            full_addr_str = driver.find_element(By.XPATH, "//h1[starts-with(text(), 'О доме')]").text
            addr_line = full_addr_str.split(' в ')[0]
            # addr_line = re.sub(' в Москве', '',full_addr_str)
            # addr_line = re.sub(' в Зеленограде', '', addr_line)
            # addr_line = re.sub(' в Зеленограде', '', addr_line)
            addr_line = re.sub(' ЗЕЛЕНОГРАД Г.', ' г. Зеленоград ', addr_line)
            parse_addr_line = secondary_addr_normalize_for_metrodistance(addr_line)
            parse_addr_line = re.sub(' г. Зеленоград, д.', ' г. Зеленоград, к.', parse_addr_line)

            if (parse_addr_line.lower()) != addr_norm:
                print('адреса не совпадают')
                logging.info('адреса не совпадают')
                print('Ввести правильный url?')
                is_cont = input()
                if is_cont in ['Y', 'y']:
                    print('Введите правильный url')
                    new_url = input()
                    driver.get(new_url)
                    # if re.search('h_info', new_url) is not None:
                    #     addr_line = re.sub(' в Москве', '',
                    #                        driver.find_element(By.XPATH, "//h1[starts-with(text(), 'О доме')]").text)
                    #     parse_addr_line = secondary_addr_normalize_for_metrodistance(addr_line)
                    #     if (parse_addr_line.lower()) != addr_norm:
                    #         print('адреса нового url не совпадают')
                    #         logging.info('адреса нового url не совпадают')
                    #         driver.quit()
                    #         return res
                    cur_url = new_url
                else:
                    driver.quit()
                    return res
            print('адреса норм')
            logging.info('адреса норм')
            res['gkh_address'] = addr_norm

            try:
                r = requests.get(cur_url)
                res['building_page_url'] = cur_url
                soup = BeautifulSoup(r.text, 'lxml')
                # скачиваем страницу инфы о доме
                full_info = soup.find('div', class_='page__content')
                li_blocks = full_info.find_all('li', class_='fi-list__item fi-list-item')
                # и ищем характеристики этажности
                for li_bl in li_blocks:
                    try:
                        label_text = li_bl.find('span', class_='fi-list-item__label').text.strip()
                        param_value = li_bl.find('span', class_='fi-list-item__value').text.strip()
                    except:
                        continue

                    if label_text == 'Этажей всего':
                        t2 = param_value.strip('\r').strip('\n').strip('\t').split('\n')
                        t3 = t2[0].strip()
                        res['gkh_total_floors'] = t3

                    elif label_text == 'Год постройки':
                        try:
                            res['construction_year'] = int(param_value.strip('\n').strip())
                        except Exception as exc:
                            print(addr_norm, 'ошибка определения года постройки ', param_value)
                            logging.info(f'{addr_norm} ошибка определения года постройки {param_value}')
                    elif label_text == 'Округ':
                        res['adm_area'] = param_value.strip('\n').strip()
                    elif label_text == 'Район':
                        res['mun_district'] = param_value.strip('\n').strip()
                    elif label_text == 'Гео-координаты':
                        try:
                            res['geo_lat'] = float(param_value.strip('\n').split('/')[0].strip())
                            res['geo_lon'] = float(param_value.strip('\n').split('/')[1].strip())
                        except Exception as exc:
                            print(addr_norm, 'ошибка определения координат ', param_value)
                            logging.info(f'{addr_norm} ошибка определения координат {param_value}')
                    elif label_text == 'Перекрытия':
                        res['overlap_material'] = param_value.strip('\n').strip()
                    elif label_text == 'Каркас':
                        res['skeleton'] = param_value.strip('\n').strip()
                    elif label_text == 'Стены':
                        res['wall_material'] = param_value.strip('\n').strip()
                    elif label_text == 'Категория':
                        res['category'] = param_value.strip('\n').strip()
                    elif label_text == 'Лифтов в подъезде' or \
                            label_text == 'Пассажирских лифтов в подъезде':
                        res['passenger_elevators_qty'] = param_value.strip('\n').strip()
                    elif label_text == 'Состояние':
                        res['condition'] = param_value.strip('\n').strip()
                    elif label_text == 'Кадастровый номер дома':
                        res['gkh_cadastr_num'] = param_value.strip('\n').strip()
                    elif label_text == 'Высота потолков':
                        try:
                            res['ceiling_height'] = int(param_value.strip('\n').strip().split(' ')[0])
                        except Exception as exc:
                            print('Ошибка определения высоты потолков - ', param_value)
                    elif label_text == 'Код адреса КЛАДР':
                        res['code_KLADR'] = param_value.strip('\n').strip()
                    elif label_text == 'Расселение по реновации':
                        if not re.search('не включен', param_value.lower()):
                            res['is_renov'] = 1
                            value_lst = param_value.strip('\n').strip().split(' ')
                            res['renov_period'] = value_lst[-4] + '-' + value_lst[-2]
                    elif label_text == 'Проживает':
                        try:
                            res['residents_qty'] = int(param_value.strip('\n').strip().split(' ')[0])
                        except Exception as exc:
                            print(addr_norm, 'ошибка определения количества проживающих ', param_value)
                            logging.info(f'{addr_norm} ошибка определения количества проживающих {param_value}')
                    else:
                        if re.search('ремонт', label_text):
                            print(addr_norm, 'упоминание капремонта  ', label_text, param_value)
                            logging.info(f'{addr_norm} упоминание капремонта {label_text} {param_value}')

                try:
                    transport_data = full_info.find('ul', class_='fi-list underground')
                    transp_bl = transport_data.find_all('li', class_='fi-list__item fi-list-item')
                    for bl in transp_bl:
                        is_walking_distance = bl.find('svg',
                                                    class_='location__how-label icon-svg').contents[1]
                        if re.search('#walking', str(is_walking_distance)):
                            res['gkh_metro_station'] = bl.find('span',
                                                               class_='fi-list-item__label').text.strip()
                            res['metro_min'] = int(bl.find('span',
                                                           class_='location__time').text
                                            .strip().split('\xa0')[0])
                            res['metro_km'] = round(float(bl.find('span',
                                                                  class_='fi-list-item__value')
                                            .text.strip().split(' ')[-2])/1000, 2)
                            break
                        print()
                except Exception as exc:
                    pass

            except (Exception, ):
                print('данные о доме не получены')
                logging.info('данные о доме не получены')
        driver.quit()
    except Exception as exc:
        print (exc)
    # try:
    #     driver.quit()
    # except (Exception, ):
    #     pass
    if pd.isna(res['gkh_metro_station']):
        res['gkh_metro_station'] = 'Нет'
    if pd.isna(res['is_renov']):
        res['is_renov'] = 0
    if not pd.isna(res['building_page_url']):
        update_gkh_base(addr_norm, res)
    else:
        print('Не получен ', res)
        logging.info(f'Объект {res} не получен')
    print(res)
    logging.info(f'Объект {res}')
    return res

# -
def renov_fill(ser):
    if str(ser[0]) != 'nan':
        return ser[0]
    elif ser[1] >= 2010:
        return 'новый дом'
    elif ser[1] < 2010:
        return 'нет в плане'
    else:
        return 'N/A'

# -
def max_floor(st):
    """ Функция очистки данных об этажности. Возвращает последнее число из строки и переводит в int"""
    try:
        return int(re.findall(r'\d+', st)[-1])
    except (Exception, ):
        return 0

# -
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

# -
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


def get_url_list_from_page(driver, cur_page_url, is_active):
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
    today = pd.to_datetime('today')
    for el in nfo:
        deadline = el.find('div', attrs={"class": "uid-text-small uid-text-gray"})\
                        .span.text.split(',')[0]
        deadline_date = datetime.datetime.strptime(deadline, "%d.%m.%Y").date()
        if (is_active and deadline_date>today) or (not is_active and deadline_date<=today):
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
                    driver, lot_single = parse_lot(driver, url)
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
    last_dup = lot_df[lot_df.duplicated(['lot_tag'], keep=False)].sort_values('bidding_date')
    last_dup = last_dup.drop_duplicates(subset=['lot_tag'], keep='last')
    # удаляем дублирующиеся лоты по тагу
    lot_df = lot_df.drop_duplicates(subset=['lot_tag'], keep=False, ignore_index=True)
    # и дописываем оптимальные из дубликатов к датафрейму
    lot_df = lot_df.append(last_dup).reset_index(drop=True)
    print('Количество уникальных объектов: ', len(lot_df))
    logging.info(f'Количество уникальных объектов: {len(lot_df)}')
    return lot_df

# -
def compare_with_final_df(lot_df, final_df):
    lot_df = lot_df.merge(final_df[['lot_tag', 'addr_norm']].rename(columns={'addr_norm':'is_present'}),
                          how='left', on='lot_tag')
    lot_df = lot_df[lot_df.is_present.isna()]
    lot_df = lot_df.drop(columns=['is_present'])
    return lot_df

# -
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
    gkh_df = pd.read_csv(GKH_BASE_FILENAME)
    gkh_df['gkh_address'] = gkh_df.gkh_address.apply(lambda x: str(x).lower())

    # проводим первичную нормализацию адресов датафрейма
    print('Первичная нормализация адресов датафрейма')
    logging.info('Первичная нормализация адресов датафрейма')
    norm_df.loc[:, 'addr_street'], norm_df.loc[:, 'addr_build_num'], \
        norm_df.loc[:, 'addr_floor_from_title'], norm_df.loc[:, 'addr_apart_num_from_title'] = zip(
            *norm_df.loc[:, 'addr_string'].progress_apply(primary_addr_normalize))

    # добавляем колонки с нормализованным адресом, и приведенным к нижнему регистру
    # по приведенному к нижнему регистру адресу пытаемся сджойнить с базой ЖКХ.
    norm_df['addr_norm'] = norm_df['addr_street'] + norm_df['addr_build_num']
    norm_df['addr_norm_lower'] = norm_df['addr_norm'].apply(lambda x: str(x).lower())
    norm_df = norm_df.merge(gkh_df, how='left', left_on='addr_norm_lower', right_on='gkh_address')

    # выделяем из базы строки, для которых не нашлось соответствия в базе ЖКХ
    quest_df = norm_df[norm_df.gkh_address.isna()]
    norm_df = norm_df.drop(norm_df[norm_df.gkh_address.isna()].index)

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
        gkh_df = pd.read_csv(GKH_BASE_FILENAME)
        gkh_df['gkh_address'] = gkh_df.gkh_address.apply(lambda x: str(x).lower())
        # убираем из нее ранее добавленные данные из ЖКХ (чтобы потом повторно попробовать сджойнить)
        quest_df = quest_df.drop(columns=GKH_FIELDS)
        # проводим повторную нормализацию адресов датафрейма
        quest_df.loc[:, 'addr_street'], quest_df.loc[:, 'addr_build_num'], \
            quest_df.loc[:, 'addr_floor_from_title'], quest_df.loc[:, 'addr_apart_num_from_title'] \
            = zip(*quest_df.loc[:, 'addr_string'].progress_apply(secondary_addr_normalize))
        # добавляем колонки с нормализованным адресом, и приведенным к нижнему регистру
        # по приведенному к нижнему регистру адресу пытаемся сджойнить с базой ЖКХ.
        quest_df['addr_norm'] = quest_df['addr_street'] + quest_df['addr_build_num']
        quest_df['addr_norm_lower'] = quest_df['addr_norm'].apply(lambda x: str(x).lower())
        quest_df = quest_df.merge(gkh_df, how='left', left_on='addr_norm_lower', right_on='gkh_address')
        # те строки, которые получилось сджойнить, добавляем в norm_df
        norm_df = norm_df.append(quest_df.query('addr_norm_lower == gkh_address'))
        quest_df = quest_df.query('addr_norm_lower != gkh_address')
        if len(quest_df) == 0:
            break
        else:
            print('Нераспознанных адресов: ', len(quest_df))
            logging.info(f'Нераспознанных адресов: {len(quest_df)}')
            # новая версия - выдает на печать только по одной паре реальный адрес / распознанный
            quest_df = quest_df.reset_index(drop=True)
            quest_df_sample = quest_df.drop_duplicates(subset=['addr_norm_lower']).copy()
            for i, row in quest_df_sample.iterrows():
                try:
                    print(row['addr_string'])
                    print(row['addr_norm_lower'])
                    print('Если нормализованный адрес совпадает с адресом в строке объявления - нажмите Y')
                    is_cont = input()
                    while is_cont not in ['Y', 'y', 'N', 'n']:
                        print('Некорректный ввод. Повторите выбор (Y/N)')
                        is_cont = input()
                    if is_cont in ['Y', 'y']:
                        new_addr_dict = {'gkh_address': row['addr_norm_lower']}
                        gkh_df = pd.concat([gkh_df, pd.DataFrame.from_records([new_addr_dict])] ,ignore_index=True)
                except Exception as e:
                    logging.info(f'Ошибка: {traceback.format_exc()}')
            gkh_df.to_csv(GKH_BASE_FILENAME, index=False)


            # TODO: откорректировать поиск нераспознанных
#                add_unrecognized()
            quest_df = quest_df.drop(columns=GKH_FIELDS)
            quest_df['addr_norm_lower'] = quest_df['addr_norm'].apply(lambda x: str(x).lower())
            quest_df = quest_df.merge(gkh_df, how='left', left_on='addr_norm_lower', right_on='gkh_address')
            norm_df = norm_df.append(quest_df.query('addr_norm_lower == gkh_address'))
            quest_df = quest_df.query('addr_norm_lower != gkh_address')
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

    return norm_df

# -
def is_metro_and_floor_data_complete(norm_df):
    """ Функция проверки полноты данных о привязке к метро и этажности зданий. При отсутствии данных в базе ЖКХ -
        запускается попытка скачать с сайта FlatInfo.ru, при невозможности - дает возможность доплнения данных
        базы ЖКХ вручную (через new_buildings.xlsx)
        :param norm_df: полный датафрейм после парсинга лотов
        :return: norm_df: датафрейм с базой ЖКХ и нормализованными адресами.
    """
    print('Проверка наличия данных о метро и этажности зданий')
    logging.info('Проверка наличия данных о метро и этажности зданий')
    out_of_data = norm_df[norm_df.gkh_metro_station.isna() | norm_df.gkh_total_floors.isna() | norm_df.construction_year.isna()] \
        .drop_duplicates(subset=['addr_norm'], keep='last')
    if out_of_data.shape[0] != 0:
        print('Отсутвуют данные по ', len(out_of_data), ' объектам')
        logging.info(f'Отсутвуют данные по {len(out_of_data)} объектам')
        print('Дождитесь окончания сбора и сохранения данных')
        out_of_data.progress_apply(lambda x: metro_and_floor_data(x.gkh_address), axis=1)
        gkh_df.to_csv(GKH_BASE_FILENAME, index=False)
        print('Данные сохранены')
        logging.info('Данные сохранены')
        norm_df = norm_df.drop(columns=GKH_FIELDS)
        norm_df = norm_df.merge(gkh_df, how='left', left_on='addr_norm_lower', right_on='gkh_address')
        out_of_data = norm_df[
            norm_df.gkh_metro_station.isna() | norm_df.gkh_total_floors.isna() | norm_df.construction_year.isna()] \
            .drop_duplicates(subset=['addr_norm'], keep='last')
    if out_of_data.shape[0] != 0:
        out_of_data.progress_apply(lambda x: metro_and_floor_data(x.gkh_address), axis=1)
        gkh_df.to_csv(GKH_BASE_FILENAME, index=False)
        norm_df = norm_df.drop(columns=GKH_FIELDS)
        norm_df = norm_df.merge(gkh_df, how='left', left_on='addr_norm_lower', right_on='gkh_address')
    return norm_df

# -
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

# -
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

# -
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

# -
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

# -
def copy_cell(src_sheet, src_row, src_col,
              tgt_sheet, tgt_row, tgt_col,
              copy_style=True):
    """Функция копирования свойств ячейки с одного листа на другой вместе с содержимым """
    cell = src_sheet.cell(src_row, src_col)
    new_cell = tgt_sheet.cell(tgt_row, tgt_col, cell.value)
    if cell.has_style and copy_style:
        new_cell._style = copy(cell._style)

# -
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

# -
def historical_processing():

    # считываем данные о прошедших торгах из файла
#    final_df = pd.read_csv(LOT_FILENAME_HISTORICAL)
    driver = start_browser_for_parse()
    driver.implicitly_wait(30)
    # url до 26.04.2023 'https://investmoscow.ru/tenders?pageNumber=1&pageSize=100&orderBy=CreateDate&orderAsc=false&objectTypes=7&tenderTypes=13&tenderStatuses=1&tradeForms=45001'
    main_page_url = 'https://investmoscow.ru/tenders?pageNumber=1&pageSize=10&orderBy=RequestEndDate&orderAsc=true&objectTypes=nsi:41:30011568&objectKinds=nsi:tender_type_portal:13&tenderStatus=nsi:tender_status_tender_filter:2&timeToPublicTransportStop.noMatter=true'
    driver.get(main_page_url)
    wait = WebDriverWait(driver, 30)
    print('Номер страницы')
    num = int(input())
    cur_page_url = 'https://investmoscow.ru/tenders?pageNumber=' + str(num) + '&pageSize=100&' + \
                   'orderBy=RequestEndDate&orderAsc=true&objectTypes=nsi:41:30011568&' + \
                   'objectKinds=nsi:tender_type_portal:13&tenderStatus=nsi:tender_status_tender_filter:2' + \
                   '&timeToPublicTransportStop.noMatter=true'

    driver, url_page_lst = get_url_list_from_page(driver, cur_page_url, is_active=False)


    unreadable_urls = []
    print('Обработка страницы ', num, 'прошедших торгов')
    logging.info(f'Обработка страницы {num} прошедших торгов ')
    for url in tqdm(url_page_lst, position=0):
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
    global gkh_df
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
        driver, url_page_lst = get_url_list_from_page(driver, cur_page_url, is_active=True)
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
            driver, lot_single = parse_lot(driver, url)
            if not (pd.isna(lot_single['start_price'])):
                lot_df = pd.concat([lot_df, lot_single.to_frame().T], ignore_index = True)
            else:
                print(f'Нечитаемая ссылка N{error_url_cnt} {url}')
                print('lot_tag ' + lot_single['lot_tag'])
                logging.info(f'Нечитаемая ссылка N{error_url_cnt} {url}')
                logging.info(f"lot_tag  {lot_single['lot_tag']}")
                unreadable_urls.append(url)
                error_url_cnt += 1
                continue
        except Exception as e:
            print(f'Нечитаемая ссылка N{error_url_cnt} {url}')
            print('lot_tag ' + lot_single['lot_tag'])
            logging.info(f'Нечитаемая ссылка N{error_url_cnt} {url}')
            logging.info(f"lot_tag  {lot_single['lot_tag']}")
            unreadable_urls.append(url)
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
    lot_df = lot_df[lot_df['finish_application_date'] > today]
    lot_df.to_csv('time_lot_df.csv',index=False)
    # lot_df = pd.read_csv('time_lot_df.csv') #, nrows = 90


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

    norm_df.to_csv(LOT_FILENAME)




# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    #
    # test_url_list=['https://torgi.gov.ru/new/public/notices/view/21000005000000010042']
    # driver = start_browser_for_parse()
    # for url_ in test_url_list:
    #     driver, lot_status, final_price = is_lot_finished(driver, url_)
    #     print(lot_status, final_price)


    LOT_FIELDS = ['lot_tag', 'auct_type', 'object_type', 'cadastr_num', 'addr_string',
                  'addr_street', 'addr_build_num', 'addr_floor_from_title', 'addr_apart_num_from_title',
                  'flat_num', 'qty_rooms', 'addr_floor', 'total_floors', 'obj_square', 'start_price',
                  'start_price_m2','deposit', 'auct_step', 'auct_form', 'start_applications_date',
                  'finish_application_date', 'participant_selection_date',
                  'bidding_date', 'results_date', 'roseltorg_url', 'torgi_url',
                  'metro_station', 'inf_sales', 'inf_food_service', 'inf_education',
                  'inf_cult_and_sport', 'inf_consumer_services', 'inf_health_care',
                  'documentation_expl', 'documentation_photo', 'investmoscow_url']
    LOT_FILENAME = r'..\realty_model_actual_lot_data.csv'
    LOT_FILENAME_HISTORICAL = r'..\realty_model_historical_lot_data.csv'
    GKH_BASE_FILENAME = r'gkh_base.csv'
    GKH_FIELDS = ['adm_area', 'mun_district', 'gkh_address', 'gkh_total_floors', 'geo_lat',
                  'geo_lon', 'overlap_material', 'skeleton', 'wall_material',
                  'category', 'residents_qty', 'construction_year', 'ceiling_height',
                  'passenger_elevators_qty', 'condition', 'gkh_metro_station',
                  'metro_km', 'metro_min', 'gkh_cadastr_num', 'code_KLADR', 'global_repair_date',
                  'is_renov', 'renov_period',
                  'building_page_url']

    SCROLL_PAUSE_TIME = 0.5
    current_date = datetime.datetime.now()
    current_date_string = current_date.strftime('%y_%m_%d_%H_%M')
    historical_processing()
    print(current_date_string)
    logging.basicConfig(
        level=logging.INFO,
        filename="model_log_" + current_date_string + ".log",
        format="%(asctime)s - %(module)s - %(levelname)s - %(funcName)s: %(lineno)d - %(message)s",
        datefmt='%H:%M:%S',
    )

    logging.info('Hello')
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

