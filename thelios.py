from email.mime import image
import os
from re import L
import sys
import json
from time import sleep
from urllib.parse import quote
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
# from datetime import datetime
import chromedriver_autoinstaller
from models.product import Product
from models.variant import Variant
from models.metafields import Metafields
import glob
import requests
# import pandas as pd

from openpyxl import Workbook
from openpyxl.drawing.image import Image as Imag
from openpyxl.utils import get_column_letter
from PIL import Image
# from natsort 

class Thelios_Scraper:
    def __init__(self, DEBUG: bool, result_filename: str) -> None:
        self.DEBUG = DEBUG
        self.result_filename = result_filename
        self.chrome_options = Options()
        self.chrome_options.add_argument('--disable-infobars')
        self.chrome_options.add_argument("--start-maximized")
        self.chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
        self.args = ["hide_console", ]
        self.browser = webdriver.Chrome(options=self.chrome_options, service_args=self.args)
        self.data = []
        pass

    def controller(self, brands: list[dict], url: str, username: str, password: str):
        try:
            self.browser.get(url)
            self.wait_until_browsing()

            if self.login(username, password):

                if self.wait_until_element_found(20, 'css_selector', 'div[data-label="Sun"] > a'):
                    print('Scraping products for')
                    for brand in brands:

                        glasses_types = []

                        if bool(brand['glasses_type']['sunglasses']): glasses_types.append('Sunglasses')
                        if bool(brand['glasses_type']['eyeglasses']): glasses_types.append('Eyeglasses')

                        for index, glasses_type in enumerate(glasses_types):
                            print(f'Brand: {brand["brand"]} | Type: {glasses_type}')
                            scraped_products = 0
                            brand_url = ''
                            
                            if glasses_type == 'Sunglasses':
                                if str('Celine').strip().lower() == str(brand['brand']).strip().lower():
                                    brand_url = 'https://my.thelios.com/it/it/Maison/c/00?q=%3Arelevance%3Atype%3ASole%3AfittingDescription%3AUniversal%3AfittingDescription%3AInternational%3Apurchasable%3Apurchasable%3AimShip%3Afalse%3AnewArrivals%3Afalse%3AallCategoriesForName%3ACeline&text=&newArrivals=false#'
                                elif str('Dior').strip().lower() == str(brand['brand']).strip().lower():
                                    brand_url = 'https://my.thelios.com/it/it/Maison/c/00?q=%3Arelevance%3Atype%3ASole%3AfittingDescription%3AUniversal%3AfittingDescription%3AInternational%3Apurchasable%3Apurchasable%3AimShip%3Afalse%3AnewArrivals%3Afalse%3AallCategoriesForName%3ADior&text=&newArrivals=false#'
                                elif str('Fendi').strip().lower() == str(brand['brand']).strip().lower():
                                    brand_url = 'https://my.thelios.com/it/it/Maison/c/00?q=%3Arelevance%3Atype%3ASole%3AfittingDescription%3AUniversal%3AfittingDescription%3AInternational%3Apurchasable%3Apurchasable%3AimShip%3Afalse%3AnewArrivals%3Afalse%3AallCategoriesForName%3AFendi&text=&newArrivals=false#'
                                elif str('Givenchy').strip().lower() == str(brand['brand']).strip().lower():
                                    brand_url = 'https://my.thelios.com/it/it/Maison/c/00?q=%3Arelevance%3Atype%3ASole%3AfittingDescription%3AUniversal%3AfittingDescription%3AInternational%3Apurchasable%3Apurchasable%3AimShip%3Afalse%3AnewArrivals%3Afalse%3AallCategoriesForName%3AGivenchy&text=&newArrivals=false#'
                                elif str('Stella McCartney').strip().lower() == str(brand['brand']).strip().lower():
                                    brand_url = 'https://my.thelios.com/it/it/Maison/c/00?q=%3Arelevance%3Atype%3ASole%3AfittingDescription%3AUniversal%3AfittingDescription%3AInternational%3Apurchasable%3Apurchasable%3AimShip%3Afalse%3AnewArrivals%3Afalse%3AallCategoriesForName%3AStella%2BMcCartney&text=&newArrivals=false#'
                            elif glasses_type == 'Eyeglasses':
                                if str('Celine').strip().lower() == str(brand['brand']).strip().lower():
                                    brand_url = 'https://my.thelios.com/it/it/Maison/c/00?q=%3Arelevance%3Atype%3AVista%3AfittingDescription%3AUniversal%3AfittingDescription%3AInternational%3Apurchasable%3Apurchasable%3AimShip%3Afalse%3AnewArrivals%3Afalse%3AallCategoriesForName%3ACeline&text=&newArrivals=false#'
                                elif str('Dior').strip().lower() == str(brand['brand']).strip().lower():
                                    brand_url = 'https://my.thelios.com/it/it/Maison/c/00?q=%3Arelevance%3Atype%3AVista%3AfittingDescription%3AUniversal%3AfittingDescription%3AInternational%3Apurchasable%3Apurchasable%3AimShip%3Afalse%3AnewArrivals%3Afalse%3AallCategoriesForName%3ADior&text=&newArrivals=false#'
                                elif str('Fendi').strip().lower() == str(brand['brand']).strip().lower():
                                    brand_url = 'https://my.thelios.com/it/it/Maison/c/00?q=%3Arelevance%3Atype%3AVista%3AfittingDescription%3AUniversal%3AfittingDescription%3AInternational%3Apurchasable%3Apurchasable%3AimShip%3Afalse%3AnewArrivals%3Afalse%3AallCategoriesForName%3AFendi&text=&newArrivals=false#'
                                elif str('Givenchy').strip().lower() == str(brand['brand']).strip().lower():
                                    brand_url = 'https://my.thelios.com/it/it/Maison/c/00?q=%3Arelevance%3Atype%3AVista%3AfittingDescription%3AUniversal%3AfittingDescription%3AInternational%3Apurchasable%3Apurchasable%3AimShip%3Afalse%3AnewArrivals%3Afalse%3AallCategoriesForName%3AGivenchy&text=&newArrivals=false#'
                                elif str('Stella McCartney').strip().lower() == str(brand['brand']).strip().lower():
                                    brand_url = 'https://my.thelios.com/it/it/Maison/c/00?q=%3Arelevance%3Atype%3AVista%3AfittingDescription%3AUniversal%3AfittingDescription%3AInternational%3Apurchasable%3Apurchasable%3AimShip%3Afalse%3AnewArrivals%3Afalse%3AallCategoriesForName%3AStella%2BMcCartney&text=&newArrivals=false#'

                            self.browser.get(brand_url)
                            self.wait_until_browsing()

                            while True:

                                self.wait_until_element_found(30, 'css_selector', 'div[class="product__listing product__grid col-xs-12 "] > div')

                                for div_tag in self.browser.find_elements(By.CSS_SELECTOR, 'div[class="product__listing product__grid col-xs-12 "] > div'):
                                    try:
                                        product_urls = []
                                        

                                        frame_code = str(div_tag.get_attribute('onclick')).strip().split("'")[-2].strip()
                                        product_url = str(div_tag.find_element(By.CSS_SELECTOR, 'a[class*="thumb primary-image"]').get_attribute('href'))
                                        variations = str(div_tag.find_element(By.CSS_SELECTOR, 'div[class="details counter-variant"]').text).strip().replace('colori', '').strip()
                                        scraped_products += 1
                                        print(scraped_products, frame_code, variations)

                                        self.open_new_tab(product_url)
                                        self.wait_until_element_found(30, 'css_selector', 'ul[class="row list-group list-group-horizontal"] > li')

                                        for _ in range(0, 30):
                                            try:
                                                for li_tag in self.browser.find_elements(By.CSS_SELECTOR, 'ul[class="row list-group list-group-horizontal"] > li'):
                                                    product_url = li_tag.find_element(By.TAG_NAME, 'a').get_attribute('href')
                                                    if 'https://my.thelios.com' not in product_url: 
                                                        product_url = f'https://my.thelios.com{product_url}'

                                                    if product_url not in product_urls: product_urls.append(product_url)
                                                break
                                            except: sleep(0.3)

                                        for index2, product_url in enumerate(product_urls):
                                            try:
                                                
                                                if product_url != self.browser.current_url:
                                                    self.browser.get(product_url)
                                                    self.wait_until_browsing()
                                                    self.wait_until_element_found(30, 'css_selector', 'div[class*="product-details name-product"]')

                                                product = Product()
                                                product.brand = brand['brand']
                                                product.url = self.browser.current_url
                                                product.type = glasses_type

                                                try: product.number = str(self.browser.find_element(By.CSS_SELECTOR, 'div[class*="product-details name-product"]').text).strip().replace('\n', ' ')
                                                except: pass
                                                product.frame_code = frame_code
                                                try: product.lens_code = str(self.browser.find_element(By.CSS_SELECTOR, 'span[class="productColour-pdp"]').text).strip()
                                                except: pass
                                                try:
                                                    text = str(self.browser.find_element(By.CSS_SELECTOR, 'div[class$="landscape-pdp-space"] > div').text).strip()
                                                    product.frame_color = str(text).split(',')[0].strip().title()
                                                    product.lens_color = str(text).split(',')[-1].strip().title()
                                                except: pass
                                                if not str(product.number): product.number = f'{product.frame_code} {product.lens_code}'

                                                variant = Variant()
                                                try: variant.title = str(self.browser.find_element(By.CSS_SELECTOR, 'div[class="variant-selector"] > div[class*="col-md-12"] > a > button').text).strip()
                                                except: pass
                                                try: variant.sku = f'{product.number} {variant.title}'
                                                except: pass
                                                variant.found_status = 1
                                                try:
                                                    p_tag_class = self.browser.find_element(By.CSS_SELECTOR, 'p[class*="stock-status"]').get_attribute('class')
                                                    if 'instock' in p_tag_class: variant.inventory_quantity = 1
                                                    else: variant.inventory_quantity = 0
                                                except: pass

                                                try:
                                                    if len(self.browser.find_elements(By.CSS_SELECTOR, 'img[class="infobox-price product-price"]')) == 2:
                                                        self.browser.execute_script("document.getElementsByClassName('price-select-container')[1].style.display = 'block'")
                                                        sleep(0.2)
                                                        for li_tag_price in self.browser.find_elements(By.CSS_SELECTOR, 'div[id="priceTypeSelect"] > ul > li'):
                                                            if str('Prezzo Confidenziale').strip().lower() == str(li_tag_price.text).strip().lower():
                                                                ActionChains(self.browser).move_to_element(li_tag_price).click().perform()
                                                                sleep(0.3)
                                                                self.wait_until_browsing()
                                                                break
                                                        for _ in range(0, 20):
                                                            try:
                                                                for price_div_tag in self.browser.find_elements(By.CSS_SELECTOR, 'div[class="price-box"]'):
                                                                    variant.wholesale_price = str(price_div_tag.text).strip().replace('€', '').strip()
                                                                    if variant.wholesale_price:
                                                                        if variant.wholesale_price[-3:] == ',00': 
                                                                            variant.wholesale_price = f'{variant.wholesale_price[:-3]}.00'
                                                                        break
                                                            except: sleep(0.3)
                                                except: pass

                                                try:
                                                    if len(self.browser.find_elements(By.CSS_SELECTOR, 'img[class="infobox-price product-price"]')) == 2:
                                                        self.browser.execute_script("document.getElementsByClassName('price-select-container')[1].style.display = 'block'")
                                                        sleep(0.2)
                                                        for li_tag_price in self.browser.find_elements(By.CSS_SELECTOR, 'div[id="priceTypeSelect"] > ul > li'):
                                                            if str('Prezzo al pubblico').strip().lower() == str(li_tag_price.text).strip().lower():
                                                                ActionChains(self.browser).move_to_element(li_tag_price).click().perform()
                                                                sleep(0.3)
                                                                self.wait_until_browsing()
                                                                break
                                                        for _ in range(0, 20):
                                                            try:
                                                                for price_div_tag in self.browser.find_elements(By.CSS_SELECTOR, 'div[class="price-box"]'):
                                                                    variant.listing_price = str(price_div_tag.text).strip().replace('€', '').strip()
                                                                    if variant.listing_price:
                                                                        if variant.listing_price[-3:] == ',00': 
                                                                            variant.listing_price = f'{variant.listing_price[:-3]}.00'
                                                                        break
                                                            except: sleep(0.3)
                                                except: pass
                                                
                                                metafields = Metafields()
                                                try: 
                                                    metafields.img_url = str(self.browser.find_element(By.CSS_SELECTOR, 'div[class="zoomImg"] > img').get_attribute('src')).strip()
                                                    if 'https://my.thelios.com' not in metafields.img_url: metafields.img_url = f'https://my.thelios.com{metafields.img_url}' 
                                                except: pass

                                                product.metafields = metafields
                                                product.variants = variant
                                                self.data.append(product)

                                                self.save_to_json(self.data)
                                            except Exception as e:
                                                if self.DEBUG: print(f'Exception in variant loop: {e}')
                                                else: pass


                                        self.close_last_tab()
                                    except Exception as e:
                                        if self.DEBUG: print(f'Exception in product loop: {e}')
                                        else: pass

                                next_page_li = self.browser.find_element(By.CSS_SELECTOR, 'li[class*="pagination-next"]')
                                if next_page_li and 'disabled' not in next_page_li.get_attribute('class'):
                                    href = next_page_li.find_element(By.TAG_NAME, 'a').get_attribute('href')
                                    if 'https://my.thelios.com' not in href: href = f'https://my.thelios.com{href}'
                                    self.browser.get(href)
                                    self.wait_until_browsing()
                                else: break 

            else: print(f'Failed to login \nURL: {self.URL}\nUsername: {str(username)}\nPassword: {str(password)}')
        except Exception as e:
            if self.DEBUG: print(f'Exception in scraper controller: {e}')
            else: pass
        finally: 
            self.browser.quit()

    def wait_until_browsing(self) -> None:
        while True:
            try:
                state = self.browser.execute_script('return document.readyState; ')
                if 'complete' == state: break
                else: sleep(0.2)
            except: pass

    def login(self, username: str, password: str) -> bool:
        login_flag = False
        try:
            if self.wait_until_element_found(20, 'xpath', '//input[@id="j_username"]'):
                self.browser.find_element(By.XPATH, '//input[@id="j_username"]').send_keys(username)
                self.browser.find_element(By.XPATH, '//input[@id="j_password"]').send_keys(password)
                try:
                    button = WebDriverWait(self.browser, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[class*="btn-block"]')))
                    button.click()

                    WebDriverWait(self.browser, 30).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'div[data-label="Sun"] > a')))
                    login_flag = True
                except Exception as e: print(e)
        except Exception as e:
            if self.DEBUG: print(f'Exception in login: {str(e)}')
            else: pass
        finally: return login_flag

    def wait_until_element_found(self, wait_value: int, type: str, value: str) -> bool:
        flag = False
        try:
            if type == 'id':
                WebDriverWait(self.browser, wait_value).until(EC.presence_of_element_located((By.ID, value)))
                flag = True
            elif type == 'xpath':
                WebDriverWait(self.browser, wait_value).until(EC.presence_of_element_located((By.XPATH, value)))
                flag = True
            elif type == 'css_selector':
                WebDriverWait(self.browser, wait_value).until(EC.presence_of_element_located((By.CSS_SELECTOR, value)))
                flag = True
            elif type == 'class_name':
                WebDriverWait(self.browser, wait_value).until(EC.presence_of_element_located((By.CLASS_NAME, value)))
                flag = True
            elif type == 'tag_name':
                WebDriverWait(self.browser, wait_value).until(EC.presence_of_element_located((By.TAG_NAME, value)))
                flag = True
        except: pass
        finally: return flag

    # def get_brand_urls(self, brand: dict) -> str:
    #     brand_urls = []
    #     try:
    #         div_tags = self.browser.find_element(By.XPATH, '//div[@id="mCSB_1_container"]').find_elements(By.XPATH, './/div[@class="brand-box col-2"]')
    #         for div_tag in div_tags:
    #             if bool(brand['glasses_type']['sunglasses']):
    #                 href = div_tag.find_element(By.XPATH, ".//a[contains(text(), 'Sun')]").get_attribute('href')
    #                 if f'codeLine1={str(brand["code"]).strip().upper()}' in href:
    #                     brand_urls.append([f'{href}&limit=80', 'Sunglasses'])
    #             if bool(brand['glasses_type']['eyeglasses']):
    #                 href = div_tag.find_element(By.XPATH, ".//a[contains(text(), 'Optical')]").get_attribute('href')
    #                 if f'codeLine1={str(brand["code"]).strip().upper()}' in href:
    #                     brand_urls.append([f'{href}&limit=80', 'Eyeglasses'])
    #     except Exception as e:
    #         if self.DEBUG: print(f'Exception in get_brand_url: {str(e)}')
    #         else: pass
    #     finally: return brand_urls

    def open_new_tab(self, url: str) -> None:
        # open category in new tab
        self.browser.execute_script('window.open("'+str(url)+'","_blank");')
        self.browser.switch_to.window(self.browser.window_handles[len(self.browser.window_handles) - 1])
        self.wait_until_browsing()
    
    def close_last_tab(self) -> None:
        self.browser.close()
        self.browser.switch_to.window(self.browser.window_handles[len(self.browser.window_handles) - 1])

    # def get_all_products_from_page(self) -> list[dict]:
    #     products_on_page = []
    #     try:
    #         for _ in range(0, 30):
    #             products_on_page = []
    #             try:
    #                 for div_tag in self.browser.find_elements(By.XPATH, '//div[@class="row mt-4 list grid-divider"]/div'): 
    #                     ActionChains(self.browser).move_to_element(div_tag).perform()
    #                     product_url, product_number, product_brand, product_gender = '', '', '', ''
    #                     sizes = []

    #                     product_url = div_tag.find_element(By.TAG_NAME, 'a').get_attribute('href')
    #                     text = str(div_tag.find_element(By.XPATH, './/p[@class="model-name"]').text).strip()
    #                     product_number = str(text.split(' ')[0]).strip()
    #                     product_name = str(text).replace(product_number, '').strip()
    #                     product_brand = str(div_tag.find_element(By.XPATH, '//div[@class="line-name d-flex justify-content-between"]/p').text).strip()
    #                     # if str('WEB').strip().lower() == str(product_brand).strip().lower():
    #                     #     try:
    #                     #         if str(product_number[:2]).strip().lower() != str('WB').strip().lower():
    #                     #             new_number = quote(product_number)
    #                     #             new_url = f'https://digitalhub.marcolin.com/shop/products?searchText={new_number}'
    #                     #             new_brand_name = self.search_for_brand_name(new_url, product_number)
    #                     #             if new_brand_name and str(new_brand_name).strip().lower() != str(product_brand).strip().lower(): product_brand = new_brand_name
    #                     #     except Exception as e: 
    #                     #         if self.DEBUG: print(f'Exception as in getting new brand name: {str(e)}')
    #                     #         else: pass

    #                     try: product_gender = str(div_tag.find_element(By.XPATH, './/div[@class="info"]/p[contains(text(), "Gender")]').text).replace('Gender:', '').strip()
    #                     except: pass
                        
    #                     try:
    #                         size_value = str(div_tag.find_element(By.XPATH, './/div[@class="info"]/p[contains(text(), "Size")]').text).replace('Size:', '').strip()
                            
    #                         if ',' in size_value:
    #                             for value in size_value.split(','):
    #                                 sizes.append(str(value).strip())
    #                         else: sizes.append(size_value)
    #                     except: pass
                        
    #                     json_data = {
    #                         'number': product_number,
    #                         'name': product_name,
    #                         'brand': product_brand,
    #                         'gender': product_gender,
    #                         'url': product_url,
    #                         'sizes': sizes
    #                     }
    #                     if json_data not in products_on_page: products_on_page.append(json_data)
    #                 break
    #             except: sleep(0.3)
    #     except Exception as e:
    #         if self.DEBUG: print(f'Exception in get_all_products_from_page: {str(e)}')
    #         else: pass
    #     finally: return products_on_page

    # def is_next_page(self) -> bool:
    #     next_page_flag = False
    #     try:
    #         next_span_style = self.browser.find_element(By.XPATH, '//span[@class="next"]').get_attribute('style')
    #         if ': hidden;' not in next_span_style: next_page_flag = True
    #     except Exception as e:
    #         if self.DEBUG: print(f'Exception in is_next_page: {str(e)}')
    #         else: pass
    #     finally: return next_page_flag

    # def move_to_next_page(self) -> None:
    #     try:
    #         current_page_number = str(self.browser.find_element(By.XPATH, '//span[@class="current"]').text).strip()
    #         next_page_span = self.browser.find_element(By.XPATH, '//span[@class="next"]')
    #         # ActionChains(self.browser).move_to_element(next_page_span).perform()
    #         ActionChains(self.browser).move_to_element(next_page_span).click().perform()
    #         self.wait_for_next_page_to_load(current_page_number)
    #     except Exception as e:
    #         if self.DEBUG: print(f'Exception in move_to_next_page: {str(e)}')
    #         else: pass
    
    # def wait_for_next_page_to_load(self, current_page_number: str) -> None:
    #     for _ in range(0, 100):
    #         try:
    #             next_page_number = str(self.browser.find_element(By.XPATH, '//span[@class="current"]').text).strip()
    #             if int(next_page_number) > int(current_page_number): 
    #                 for _ in range(0, 30):
    #                     try:
    #                         for div_tag in self.browser.find_elements(By.XPATH, '//div[@class="row mt-4 list grid-divider"]/div'):
    #                             div_tag.find_element(By.XPATH, './/p[@class="model-name"]').text
    #                         break
    #                     except: sleep(0.3)
    #                 break
    #         except: sleep(0.3)
     
    # def move_to_first_variant(self) -> None:
    #     while True:
    #         try:
    #             elements = self.browser.find_elements(By.XPATH, '//span[@class="n-arrow prev slick-arrow"]')
    #             if len(elements) == 2:
    #                 ActionChains(self.browser).move_to_element(elements[0]).click().perform()
    #                 sleep(0.3)
    #             else: break
    #         except: pass
    
    # def select_variant_image(self, divs: list) -> None:
    #     for _ in range(0, 5):
    #         try:
    #             ActionChains(self.browser).move_to_element(divs).click().perform()
    #             sleep(0.3)
    #             break
    #         except:
    #             try:
    #                 elements = self.browser.find_elements(By.XPATH, '//span[@class="n-arrow next slick-arrow"]')
    #                 if len(elements) == 2:
    #                     ActionChains(self.browser).move_to_element(elements[0]).click().perform()
    #                     sleep(0.3)
    #             except Exception as e: 
    #                 if self.DEBUG: print(str(e))
    #                 else: pass

    def click_to_make_price_visible(self) -> None:
        try:
            if not self.is_price_visible():
                element = self.browser.find_element(By.XPATH, '//svg-icon[@class="eye-icon"]')
                ActionChains(self.browser).move_to_element(element).click().perform()
                sleep(0.4)

                for li in self.browser.find_elements(By.XPATH, '//ul[@aria-labelledby="basic-link"]/li'):
                    if str('Cost and SRP').strip().lower() in str(li.text).strip().lower():
                        ActionChains(self.browser).move_to_element(li).click().perform()
                        self.wait_until_price_is_shown()
        except Exception as e:
            if self.DEBUG: print(f'Exception in click_to_make_price_visible: {str(e)}')
            else: pass

    def is_price_visible(self) -> bool:
        try:
            for td in self.browser.find_elements(By.XPATH, '//table[@class="table table-borderless"]/tbody/tr/td'):
                if str('Suggested Retail Price').strip().lower() in str(td.text).strip().lower(): return True
            return False
        except: return False

    def wait_until_price_is_shown(self) -> bool:
        flag = False
        for _ in range(0, 30):
            try:
                tds_label = self.browser.find_elements(By.XPATH, '//table[@class="table table-borderless"]/tbody/tr/td')
                for i in range(0, len(tds_label)):
                    if str('Suggested Retail Price').strip().lower() in str(tds_label[i].text).strip().lower(): 
                        flag = True
                        break
                if flag: break
            except: sleep(0.3)
            finally: return flag

    def get_size_price_status(self):
        size_titles, wholesale_prices, listing_prices, availability, gtin = [], [], [], [], []
        sizes = []
        caliber, rod, bridge = '', '', ''
        try:
            trs = self.browser.find_elements(By.XPATH, '//table[@class="table table-borderless"]/tbody/tr')
            
            for j in range(1, len(trs)):
                tr = trs[j].find_element(By.XPATH, '//table[@class="table table-borderless inner-table"]/tr')
                
                # tds_value = tr.find_elements_by_tag_name('td')
                try:
                    value = str(tr.find_elements(By.XPATH, ".//td[contains(text(), '€')]")[0].text).replace('€', '').strip()
                    value = f"{str(value[0:-3]).strip().replace(',', '').strip()}.00"
                    wholesale_prices.append(value)
                except Exception as e: 
                    if self.DEBUG: print(f'Exception in wholesale_price: {str(e)}')
                    else: pass

                try:
                    value = str(tr.find_elements(By.XPATH, ".//td[contains(text(), '€')]")[1].text).replace('€', '').strip()
                    value = f"{str(value[0:-3]).strip().replace(',', '').strip()}.00"
                    listing_prices.append(value)
                except Exception as e: 
                    if self.DEBUG: print(f'Exception in listing_prices: {str(e)}')
                    else: pass
                
            for tr in self.browser.find_elements(By.XPATH, '//table[@class="table table-borderless inner-table"]/tr'):
                tds_value = tr.find_elements(By.TAG_NAME, 'td')
                try:
                    caliber, rod, bridge = str(tds_value[0].text).strip(), str(tds_value[1].text).strip(), str(tds_value[2].text).strip()
                    # print(len(tds_value), f'{caliber}-{rod}-{bridge}')
                    size_titles.append(caliber)
                    sizes.append(f'{caliber}-{rod}-{bridge}')
                    # if not metafields.product_size: metafields.product_size =f'{caliber}-{rod}-{bridge}'
                    caliber, rod, bridge = '', '', ''
                except Exception as e: 
                    if self.DEBUG: print(f'Exception in product size: {str(e)}')
                    else: pass

                
                try:
                    span_tag_class = tr.find_element(By.CSS_SELECTOR, 'span[class^="availability"]').get_attribute('class')
                    # if str('a-0').strip().lower() in  str(span_tag_class).strip().lower(): availability.append('Draft')
                    # if str('a-1').strip().lower() in  str(span_tag_class).strip().lower(): availability.append('Draft')
                    if str('a-2').strip().lower() in  str(span_tag_class).strip().lower(): availability.append('Active')
                    else: availability.append('Draft')
                    # elif str('a-3').strip().lower() in  str(span_tag_class).strip().lower(): availability.append('Draft')
                except: 
                    availability.append('Not Available')

            
            # for j in range(1, len(trs)):
            try:
                for button in self.browser.find_elements(By.XPATH, '//svg-icon[@class="arrow-icon"]'):
                    ActionChains(self.browser).move_to_element(button).perform()
                    button.click()
                    sleep(0.5)
                    for _ in range(0, 20):
                        try:
                            tags = self.browser.find_elements(By.XPATH, '//table[@class="table table-borderless inner-table open-shadow"]/tr[@class="d-flex drawer"]')
                            for tag in tags:    
                                g = str(tag.find_element(By.CSS_SELECTOR, 'div[class$="ean-detail"] > p > span').text).strip()
                                if g: 
                                    if g not in gtin: gtin.append(g)
                                else: gtin.append('')
                            # close_element = self.browser.find_element(By.XPATH, '//table[@class="table table-borderless inner-table open-shadow"]/tr[@class="d-flex"]').find_element(By.XPATH, '//svg-icon[@class="arrow-icon"]')
                            # ActionChains(self.browser).move_to_element(close_element).perform()
                            # close_element.click()
                            # sleep(0.3)
                            break
                        except Exception as e:pass
            except Exception as e: pass
        except Exception as e:
            if self.DEBUG: print(f'Exception in get_size_price_status: {str(e)}')
            else: pass
        finally: return size_titles, wholesale_prices, listing_prices, availability, gtin, sizes

    def save_to_json(self, products: list[Product]):
        try:
            json_products = []
            for product in products:
                json_varinats = []
                for index, variant in enumerate(product.variants):
                    json_varinat = {
                        'position': (index + 1), 
                        'title': variant.title, 
                        'sku': variant.sku, 
                        'inventory_quantity': variant.inventory_quantity,
                        'found_status': variant.found_status,
                        'wholesale_price': variant.wholesale_price,
                        'listing_price': variant.listing_price, 
                        'barcode_or_gtin': variant.barcode_or_gtin,
                        'size': variant.size,
                        'weight': variant.weight
                    }
                    json_varinats.append(json_varinat)
                json_product = {
                    'brand': product.brand, 
                    'number': product.number, 
                    'name': product.name, 
                    'frame_code': product.frame_code, 
                    'frame_color': product.frame_color, 
                    'lens_code': product.lens_code, 
                    'lens_color': product.lens_color, 
                    'status': product.status, 
                    'type': product.type, 
                    'url': product.url, 
                    'metafields': [
                        { 'key': 'for_who', 'value': product.metafields.for_who },
                        { 'key': 'product_size', 'value': product.metafields.product_size }, 
                        { 'key': 'lens_material', 'value': product.metafields.lens_material }, 
                        { 'key': 'lens_technology', 'value': product.metafields.lens_technology }, 
                        { 'key': 'frame_material', 'value': product.metafields.frame_material }, 
                        { 'key': 'frame_shape', 'value': product.metafields.frame_shape },
                        { 'key': 'gtin1', 'value': product.metafields.gtin1 }, 
                        { 'key': 'img_url', 'value': product.metafields.img_url }
                    ],
                    'variants': json_varinats
                }
                json_products.append(json_product)
            
           
            with open(self.result_filename, 'w') as f: json.dump(json_products, f)
            
        except Exception as e:
            if self.DEBUG: print(f'Exception in save_to_json: {e}')
            else: pass


def read_data_from_json_file(DEBUG, result_filename: str):
    data = []
    try:
        files = glob.glob(result_filename)
        if files:
            f = open(files[-1])
            json_data = json.loads(f.read())
            products = []

            for json_d in json_data:
                number, frame_code, brand, img_url, frame_color, lens_color = '', '', '', '', '', ''
                # product = Product()
                brand = json_d['brand']
                number = str(json_d['number']).strip().upper()
                if '/' in number: number = number.replace('/', '-').strip()
                # product.name = str(json_d['name']).strip().upper()
                frame_code = str(json_d['frame_code']).strip().upper()
                if '/' in frame_code: frame_code = frame_code.replace('/', '-').strip()
                frame_color = str(json_d['frame_color']).strip().title()
                # product.lens_code = str(json_d['lens_code']).strip().upper()
                lens_color = str(json_d['lens_color']).strip().title()
                # product.status = str(json_d['status']).strip().lower()
                # product.type = str(json_d['type']).strip().title()
                # product.url = str(json_d['url']).strip()
                # metafields = Metafields()
                
                for json_metafiels in json_d['metafields']:
                    # if json_metafiels['key'] == 'for_who':metafields.for_who = str(json_metafiels['value']).strip().title()
                    # elif json_metafiels['key'] == 'product_size':metafields.product_size = str(json_metafiels['value']).strip().title()
                    # elif json_metafiels['key'] == 'activity':metafields.activity = str(json_metafiels['value']).strip().title()
                    # elif json_metafiels['key'] == 'lens_material':metafields.lens_material = str(json_metafiels['value']).strip().title()
                    # elif json_metafiels['key'] == 'graduabile':metafields.graduabile = str(json_metafiels['value']).strip().title()
                    # elif json_metafiels['key'] == 'interest':metafields.interest = str(json_metafiels['value']).strip().title()
                    # elif json_metafiels['key'] == 'lens_technology':metafields.lens_technology = str(json_metafiels['value']).strip().title()
                    # elif json_metafiels['key'] == 'frame_material':metafields.frame_material = str(json_metafiels['value']).strip().title()
                    # elif json_metafiels['key'] == 'frame_shape':metafields.frame_shape = str(json_metafiels['value']).strip().title()
                    # elif json_metafiels['key'] == 'gtin1':metafields.gtin1 = str(json_metafiels['value']).strip().title()
                    if json_metafiels['key'] == 'img_url':img_url = str(json_metafiels['value']).strip()
                    # elif json_metafiels['key'] == 'img_360_urls':
                    #     value = str(json_metafiels['value']).strip()
                    #     if '[' in value: value = str(value).replace('[', '').strip()
                    #     if ']' in value: value = str(value).replace(']', '').strip()
                    #     if "'" in value: value = str(value).replace("'", '').strip()
                    #     for v in value.split(','):
                    #         metafields.img_360_urls = str(v).strip()
                # product.metafields = metafields
                for json_variant in json_d['variants']:
                    sku, price = '', ''
                    # variant = Variant()
                    # variant.position = json_variant['position']
                    # variant.title = str(json_variant['title']).strip()
                    sku = str(json_variant['sku']).strip().upper()
                    if '/' in sku: sku = sku.replace('/', '-').strip()
                    # variant.inventory_quantity = json_variant['inventory_quantity']
                    # variant.found_status = json_variant['found_status']
                    wholesale_price = str(json_variant['wholesale_price']).strip()
                    listing_price = str(json_variant['listing_price']).strip()
                    # variant.barcode_or_gtin = str(json_variant['barcode_or_gtin']).strip()
                    # variant.size = str(json_variant['size']).strip()
                    # variant.weight = str(json_variant['weight']).strip()
                    # product.variants = variant

                    image_attachment = download_image(img_url)
                    if image_attachment:
                        with open(f'Images/{sku}.jpg', 'wb') as f: f.write(image_attachment)
                        crop_downloaded_image(f'Images/{sku}.jpg')
                    data.append([number, frame_code, frame_color, lens_color, brand, sku, wholesale_price, listing_price])
    except Exception as e:
        if DEBUG: print(f'Exception in read_data_from_json_file: {e}')
        else: pass
    finally: return data

def download_image(url):
    image_attachment = ''
    try:
        headers = {
            'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
            'accept-Encoding': 'gzip, deflate, br',
            'accept-Language': 'en-US,en;q=0.9',
            'cache-Control': 'max-age=0',
            'sec-ch-ua': '"Google Chrome";v="95", "Chromium";v="95", ";Not A Brand";v="99"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
            'sec-fetch-dest': 'document',
            'sec-fetch-mode': 'navigate',
            'sec-fetch-site': 'none',
            'Sec-Fetch-User': '?1',
            'upgrade-insecure-requests': '1',
        }
        counter = 0
        while True:
            try:
                response = requests.get(url=url, headers=headers, timeout=20)
                # print(response.status_code)
                if response.status_code == 200:
                    # image_attachment = base64.b64encode(response.content)
                    image_attachment = response.content
                    break
                else: print(f'{response.status_code} found for downloading image')
            except: sleep(0.3)
            counter += 1
            if counter == 10: break
    except Exception as e: print(f'Exception in download_image: {str(e)}')
    finally: return image_attachment

def crop_downloaded_image(filename):
    try:
        im = Image.open(filename)
        width, height = im.size   # Get dimensions
        new_width = 1120
        new_height = 600
        if width > new_width and height > new_height:
            left = (width - new_width)/2
            top = (height - new_height)/2
            right = (width + new_width)/2
            bottom = (height + new_height)/2
            im = im.crop((left, top, right, bottom))
            im.save(filename)
        elif height > new_height:
            left = (width - new_width)/2
            top = (height - new_height)/2
            right = (width + new_width)/2
            bottom = (height + new_height)/2
            im = im.crop((left, top, right, bottom))
            im.save(filename)
    except Exception as e: print(f'Exception in crop_downloaded_image: {e}')

def saving_picture_in_excel(data: list):
    workbook = Workbook()
    worksheet = workbook.active

    worksheet.cell(row=1, column=1, value='Model Code')
    worksheet.cell(row=1, column=2, value='Lens Code')
    worksheet.cell(row=1, column=3, value='Color Frame')
    worksheet.cell(row=1, column=4, value='Color Lens')
    worksheet.cell(row=1, column=5, value='Brand')
    worksheet.cell(row=1, column=6, value='SKU')
    worksheet.cell(row=1, column=7, value='Wholesale Price')
    worksheet.cell(row=1, column=8, value='Listing Price')
    worksheet.cell(row=1, column=9, value="Image")

    for index, d in enumerate(data):
        new_index = index + 2

        worksheet.cell(row=new_index, column=1, value=d[0])
        worksheet.cell(row=new_index, column=2, value=d[1])
        worksheet.cell(row=new_index, column=3, value=d[2])
        worksheet.cell(row=new_index, column=4, value=d[3])
        worksheet.cell(row=new_index, column=5, value=d[4])
        worksheet.cell(row=new_index, column=6, value=d[5])
        worksheet.cell(row=new_index, column=7, value=d[6])
        worksheet.cell(row=new_index, column=8, value=d[7])

        image = f'Images/{d[-3]}.jpg'
        if os.path.exists(image):
            im = Image.open(image)
            width, height = im.size
            worksheet.row_dimensions[new_index].height = height
            worksheet.add_image(Imag(image), anchor='I'+str(new_index))
            # col_letter = get_column_letter(7)
            # worksheet.column_dimensions[col_letter].width = width

    workbook.save('Thelios Results.xlsx')

DEBUG = True
try:
    pathofpyfolder = os.path.realpath(sys.argv[0])
    # get path of Exe folder
    path = pathofpyfolder.replace(pathofpyfolder.split('\\')[-1], '')
    # download chromedriver.exe with same version and get its path
    if os.path.exists('chromedriver.exe'): os.remove('chromedriver.exe')
    if os.path.exists('Thelios Results.xlsx'): os.remove('Thelios Results.xlsx')

    chromedriver_autoinstaller.install(path)
    if '.exe' in pathofpyfolder.split('\\')[-1]: DEBUG = False
    
    f = open('Thelios start.json')
    json_data = json.loads(f.read())
    f.close()

    brands = json_data['brands']

    
    f = open('requirements/thelios.json')
    data = json.loads(f.read())
    f.close()
    url = data['url']
    username = data['username']
    password = data['password']
    
    result_filename = 'requirements/Thelios Results.json'
    Thelios_Scraper(DEBUG, result_filename).controller(brands, url, username, password)
    
    for filename in glob.glob('Images/*'): os.remove(filename)
    data = read_data_from_json_file(DEBUG, result_filename)
    os.remove(result_filename)

    saving_picture_in_excel(data)

except Exception as e:
    if DEBUG: print('Exception: '+str(e))
    else: pass
