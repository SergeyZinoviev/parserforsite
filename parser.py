# !/usr/bin/python3
# -*- coding: utf-8 -*-

import xlsxwriter
import requests
from bs4 import BeautifulSoup as bs

headers = {'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9', 'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/99.0.4844.74 Safari/537.36'}
base_url = f"https://www.studiorent.ru/studios/?city=33&page="
pages = int(input('Укажите кол-во страниц для парсинга: '))

studios =[]

def stud_pars(base_url, headers):
        zero = 0
        while pages > zero:
                zero = str(zero)
                session = requests.Session()
                request = session.get(base_url + zero, headers = headers)
                print(f'Идет загрузка {int(zero)+1} страницы')
                if request.status_code == 200:
                        soup = bs(request.content, 'html.parser')
                        divs = soup.find_all("article", "s")
                        process_import = 0
                        for div in divs: 
                                #Название студии                               
                                title = div.find('small', "text-muted").text

                                #Ссылка на студию с подробной информацией и фото                                
                                href = 'https://www.studiorent.ru'+div.find('a', "sn counters-click")['href']

                                #Адрес
                                adres = div.find('div', "col-8 col-md-8 col-lg-5 mb-4")
                                adres.header.decompose()
                                adres.span.decompose()
                                adrspisok =[]
                                for i in adres:
                                    adrspisok.append(i.text)
                                adres = adrspisok[1]

                                #Метро
                                metro = adrspisok[-4]
                                #Заходим на подробную страницу с телефоном и остальным
                                result = requests.get(href)
                                content = result.text
                                soup1 = bs(content, 'lxml')
                                # Телефон
                                phone = soup1.find('div', 'col-12 col-sm-8 col-lg-5 mt-3').find('a').text

                                #Поиск почты
                                contact = soup1.find('div', 'col-12 col-sm-8 col-lg-5 mt-3')
                                contsp = []
                                for i in contact:
                                        contsp.append(i.text)
                                for i in contsp:
                                        if '@' in i:
                                                r = i.split(' ')
                                        for i in r:
                                                if '@' in i:
                                                        email = i                              
                                

                                #Минимальное время аренды
                                try:
                                        time_ot = div.find('span', attrs = {"title":"Минимальное время аренды"}).text
                                except:
                                        time_ot = ''
                                #Размер студии
                                try:
                                        size = div.find('span', "rooms").text
                                except:
                                        size = 'None'
                                #Цена до
                                second_string = []
                                text12 = div.find('div', "params_list col-8").find_all('span')
                                for i in text12:
                                    second_string.append(i.text)
                                #Цена от
                                try:
                                        text1 = div.find('span', "price").text
                                except:
                                        text1 = ''
                                try:
                                        text2 = second_string[-1]
                                except:
                                        text2 = ''
                                price = f'{text1} - {text2}'
                                #Все данные с одной студии
                                all_txt = [title, adres, metro, size, price, time_ot, href, phone, email]
                                studios.append(all_txt)
                                process_import += 1
                                # print(f'Загрузка блока {process_import} из 50')
                        zero = int(zero)
                        zero += 1
                        

                else:
                        print('error')

                # Запись в Excel файл
                workbook = xlsxwriter.Workbook('photo_studios.xlsx')
                worksheet = workbook.add_worksheet()

                # Cтили форматирования
                bold = workbook.add_format({'bold': 1})
                bold.set_align('center')
                center_H_V = workbook.add_format()
                center_H_V.set_align('center')
                center_H_V.set_align('vcenter')
                center_V = workbook.add_format()
                center_V.set_align('vcenter')
                cell_wrap = workbook.add_format()
                cell_wrap.set_text_wrap()

                # Ширина колонок
                worksheet.set_column(0, 0, 45) # A 
                worksheet.set_column(1, 1, 54) # B
                worksheet.set_column(2, 2, 27) # C
                worksheet.set_column(3, 3, 10) # D
                worksheet.set_column(4, 4, 12) # E
                worksheet.set_column(5, 5, 26) # F
                worksheet.set_column(6, 6, 45) # G
                worksheet.set_column(7, 7, 32) # H
                worksheet.set_column(8, 8, 32) # I

                worksheet.write('A1', 'Название', bold)
                worksheet.write('B1', 'Адрес', bold)
                worksheet.write('C1', 'Метро', bold)
                worksheet.write('D1', 'Размер', bold)
                worksheet.write('E1', 'Цена', bold)
                worksheet.write('F1', 'Минимум время аренды', bold)
                worksheet.write('G1', 'Ссылка', bold)
                worksheet.write('H1', 'Телефон', bold)
                worksheet.write('I1', 'Почта', bold)

                row = 1
                col = 0
                for i in studios:
                        worksheet.write_string (row, col, i[0], center_V)
                        worksheet.write_string (row, col + 1, i[1], center_H_V)
                        worksheet.write_string (row, col + 2, i[2], center_H_V)
                        worksheet.write_string (row, col + 3, i[3], center_H_V)
                        worksheet.write_string (row, col + 4, i[4], center_V)
                        worksheet.write_string (row, col + 5, i[5], center_H_V)
                        worksheet.write_url (row, col + 6, i[6])
                        worksheet.write_string (row, col + 7, i[7], center_H_V)
                        worksheet.write_string (row, col + 8, i[8], center_H_V)
                        row += 1
                print(f'Страница {zero} из {pages} загружена')
        workbook.close()

stud_pars(base_url, headers)