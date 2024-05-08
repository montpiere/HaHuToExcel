import requests
from bs4 import BeautifulSoup
import xlsxwriter
import os
import datetime
import time


car_counter = 1

filename = f'hahutoexcel-{datetime.datetime.now().strftime("%Y%m%d%H%M%S")}.xlsx'

fuel_types_array = ["Benzin", "Dízel", "Benzin/Gáz", "LPG", "CNG", "Hibrid", "Hibrid (Benzin)",
                    "Hibrid (Dízel)", "Elektromos", "Etanol", "Biodízel", "Gáz"]

# EXCEL
workbook = xlsxwriter.Workbook(filename)
worksheet = workbook.add_worksheet()

# cell formatting
worksheet.set_row(0, 21)
worksheet.set_column(0, 1, 13)
worksheet.set_column(2, 2, 97.6)
worksheet.set_column(3, 10, 18)
worksheet.set_column('G:G', 10)
worksheet.hide_gridlines(2)
worksheet.autofilter('A1:K1')
worksheet.freeze_panes(1, 0)

# styles
head_style = workbook.add_format(
    {'bold': 1, 'underline': 1, 'font_size': 11, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#e6e6e6',
     'border': 1, 'border_color': '#aaaaaa'})
kivitel_style = workbook.add_format({'border': 1, 'border_color': '#aaaaaa'})
link_style = workbook.add_format(
    {'font_color': 'blue', 'underline': 1, 'align': 'left', 'border': 1, 'border_color': '#aaaaaa'})
center_align_style = workbook.add_format({'align': 'center', 'border': 1, 'border_color': '#aaaaaa'})
km_style = workbook.add_format({'align': 'center', 'num_format': '# ##0" km"', 'border': 1, 'border_color': '#aaaaaa'})
vetelar_style = workbook.add_format(
    {'align': 'center', 'num_format': '# ##,,0 Ft', 'border': 1, 'border_color': '#aaaaaa'})
le_style = workbook.add_format({'align': 'center', 'num_format': '# ##0" LE"', 'border': 1, 'border_color': '#aaaaaa'})
kw_style = workbook.add_format({'align': 'center', 'num_format': '# ##0" KW"', 'border': 1, 'border_color': '#aaaaaa'})
kbcm_style = workbook.add_format(
    {'align': 'center', 'num_format': '# ##0" cm³"', 'border': 1, 'border_color': '#aaaaaa'})
evjarat_style = workbook.add_format(
    {'align': 'center', 'border': 1, 'border_color': '#aaaaaa'})
honap_style = workbook.add_format(
    {'align': 'center', 'border': 1, 'border_color': '#aaaaaa'})

# column name
worksheet.write('A1', 'Márka', head_style)
worksheet.write('B1', 'Modell', head_style)
worksheet.write('C1', 'Kivitel', head_style)
worksheet.write('D1', 'Vételár', head_style)
worksheet.write('E1', 'Üzemanyag', head_style)
worksheet.write('F1', 'Évjárat', head_style)
worksheet.write('G1', 'Hónap', head_style)
worksheet.write('H1', 'Futott km', head_style)
worksheet.write('I1', 'Hengerűrtartalom', head_style)
worksheet.write('J1', 'LE', head_style)
worksheet.write('K1', 'KW', head_style)


def read_urls_from_txt():
    with open('links.txt', 'r', encoding='utf-8') as file:
        lines = file.readlines()
    lines = [line.strip() for line in lines]

    return lines


def get_url_data(_url):
    def get_soup(url):
        page = requests.get(url)
        soup = BeautifulSoup(page.content, 'html.parser')
        return soup

    def get_page_numbers(_soup):
        pagination = _soup.find(class_="pagination")
        if pagination is None:
            all_page_number = 0
        else:
            all_page_number = len(pagination.text.replace("\n", "").strip())

        return all_page_number

    def get_car_data(_data):
        find_link = _data.find('h3')
        name_info = find_link.a
        name = name_info.get_text()

        marka = name.split()[0]
        modell = name.split()[1]
        kivitel = str(name[((len(marka) + len(modell)) + 2):])
        link = name_info['href']

        info_data = _data.find_all(class_="info")

        ev = '-'
        honap = '-'
        uzemanyag = '-'
        km = '-'
        kbcm = '-'
        le = '-'
        kw = '-'

        try:
            price = _data.find(class_="pricefield-primary").get_text()
        except Exception:
            price = _data.find(class_="pricefield-primary-highlighted").get_text()

        price = price[:-3].replace(u'\xa0', u'')

        try:
            price = int(price)
        except Exception:
            price = '-'

        # additional info
        for i in range(len(info_data)):
            add_info_data = info_data[i].get_text().replace(u',', u'').replace(u'\xa0', u'')

            if add_info_data in fuel_types_array:
                uzemanyag = add_info_data

            elif 'km' in add_info_data:
                km = int(add_info_data[:-2])

            elif 'cm³' in add_info_data:
                kbcm = int(add_info_data[:-3])

            elif 'LE' in add_info_data:
                le = int(add_info_data[:-2])

            elif 'kW' in add_info_data:
                kw = int(add_info_data[:-2])

            else:
                evjarat = add_info_data
                evjarat = evjarat.replace(u'/', u'')
                ev = int(evjarat[0:4])
                if len(evjarat) > 4:
                    honap = int(evjarat[4:])
                else:
                    honap = '-'

        global car_counter
        print(f"{car_counter}. {marka} {modell} {kivitel} | {price} ft | {ev}/{honap}")

        # write data to excel
        worksheet.write(car_counter, 0, marka, center_align_style)
        worksheet.write(car_counter, 1, modell, center_align_style)
        worksheet.write_url(car_counter, 2, link, link_style, string=kivitel)
        worksheet.write(car_counter, 3, price, vetelar_style)
        worksheet.write(car_counter, 4, uzemanyag, center_align_style)
        worksheet.write(car_counter, 5, ev, evjarat_style)
        worksheet.write(car_counter, 6, honap, honap_style)
        worksheet.write(car_counter, 7, km, km_style)
        worksheet.write(car_counter, 8, kbcm, kbcm_style)
        worksheet.write(car_counter, 9, le, le_style)
        worksheet.write(car_counter, 10, kw, kw_style)

        car_counter += 1

    def get_page_data(_soup):
        return _soup.find(class_="list-view").select('div[class*="row talalati-sor"]')

    soup_main_page = get_soup(_url)
    page_numbers = get_page_numbers(soup_main_page)

    cars_list = get_page_data(soup_main_page)

    for car in cars_list:
        get_car_data(car)

    if page_numbers > 0:
        for i in range(2, page_numbers+1):
            soup_sub_page = get_soup(f"{_url}/page{i}")
            cars_list = get_page_data(soup_sub_page)
            for car in cars_list:
                get_car_data(car)


def getdata():
    start_time = time.time()
    urls = read_urls_from_txt()
    for url in urls:
        get_url_data(url)
    # get_url_data(urls[0])

    # close work with Excel file
    workbook.close()
    print("")
    print("A fájl elkészült. " + filename + " - excel fájl megnyitása...")

    # open Excel file
    command = f'start excel.exe {filename}'
    os.system(command)
    print("")
    print("Running time :  %s seconds" % (time.time() - start_time))


if __name__ == '__main__':
    getdata()
