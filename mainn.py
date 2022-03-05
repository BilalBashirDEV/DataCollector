from bs4 import BeautifulSoup
import pandas as pd
from selenium import webdriver
import random
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from time import sleep
import json

main_url = 'https://safer.fmcsa.dot.gov/CompanySnapshot.aspx'

SELENIUM = r'C:\bin\chromedriver.exe'
EXCEL_PATH = 'Women-39s Sweaters and Hoodies - Order Sheet.xlsx'
table_data = []
data_json = []


def n_digit_random(len_, floor=1):
    top = 10 ** len_
    if floor > top:
        raise ValueError(f"Floor {floor} must be less than requested top {top}")
    return f'{random.randrange(floor, top):0{len_}}'


def get_values_marked_x(tr1, i, total_tr1, t_header):
    tr2 = tr1[i + 1].tbody.find_all('tr')
    tbdy = tr2[1].find_all('tbody')
    t_row = {}
    t_list1 = []
    try:
        t_header = tr1[i].find('a', {'class': 'querylabel'}).text
    except Exception as e:
        print(e)
    for tbod in tbdy:
        tbod = tbod.find_all('tr')

        j = 1
        while j < total_tr1:
            if tbod[j].find('td', {'class': 'queryfield'}).text == 'X':
                t_list1.append(tbod[j].find_all('td')[1].text)
            j = j + 1

    t_row[t_header] = t_list1
    table_data.append(t_row)


def other_tables(table1, total_tr, t_header):
    tr2 = table1.tbody.find_all('tr')
    tbdy = tr2[1].find_all('tbody')
    t_row = {}
    t_list1 = []
    for tbod in tbdy:
        tbod = tbod.find_all('tr')
        j = 1
        while j < total_tr:
            if tbod[j].find('td', {'class': 'queryfield'}).text == 'X':
                t_list1.append(tbod[j].find_all('td')[1].text)
            j = j + 1

    t_row[t_header] = t_list1
    table_data.append(t_row)


def bot(num):
    try:
        chrome_options = Options()
        # chrome_options.add_argument(r"user-data-dir=C:\Users\HP\AppData\Local\Google\Chrome\User Data\Profile 1")
        chrome_options.add_argument("--start-maximized")
        browser = webdriver.Chrome(SELENIUM, options=chrome_options)

        while num < 1321999:
            browser.get(main_url)
            try:
                element = WebDriverWait(browser, 5).until(
                    ec.presence_of_element_located((By.ID, 2))
                )
                element.click()
                value = WebDriverWait(browser, 5).until(
                    ec.presence_of_element_located((By.ID, 4))
                )
                value.send_keys(num)
            except Exception as e:
                print(e)
                continue

            sleep(1)
            try:
                username = WebDriverWait(browser, 5).until(
                    ec.presence_of_element_located((By.XPATH, '/html/body/form/p/table/tbody/tr[4]/td/input'))
                )
                username.click()
            except Exception as e:
                print(e)
                continue
            bs_parser = BeautifulSoup(browser.page_source, "html.parser")
            try:
                bs_parser.find('table', {'summary': 'Record Inactive'}).tbody.find('i').text
                print('Record Inactive')
                num += 1
                continue
            except:
                try:
                    bs_parser.find('table', {'summary': 'For formatting purpose'}).tbody.find('i').text
                    print('Record Not Found')
                    num += 1
                    continue
                except:
                    print('Found')

            table = bs_parser.find("table", {'summary': 'For formatting purpose'})
            tr = table.tbody.find_all("tr")
            i = 1
            while i < 12:
                for th, td in zip(tr[i].find_all("th"), tr[i].find_all("td")):
                    t_row = {th.text: td.text.replace('\n', '').strip()}
                    table_data.append(t_row)
                i = i + 1

            total_tr = 5
            header = tr[12].find('a', {'class': 'querylabel'}).text
            # get_values_marked_x(tr[13], i, total_tr, header)
            other_tables(tr[13], total_tr, header)

            table1 = table.find("table", {'summary': 'Carrier Operation'})
            total_tr = 2
            header = 'Carrier Operation'
            other_tables(table1, total_tr, header)

            table1 = table.find("table", {'summary': 'Cargo Carried'})
            total_tr = 11
            header = 'Cargo Carried'
            other_tables(table1, total_tr, header)
            dictionary = {'MXCode': num}
            for dat in table_data:
                for key, value in dat.items():
                    dictionary[key] = value

            dictionary['Called At'] = ''
            dictionary['Comments'] = ''
            dictionary['Notes'] = ''

            data_json.append(dictionary)
            sleep(1)
            num += 1
            browser.close()
    except:
        print('Bad happened')
        bot(num)

    with open('result.json', 'w') as fp:
        json.dump(data_json, fp, indent=4)


if __name__ == "__main__":
    n = 1321000
    bot(n)
    df = pd.read_json('result.json')
    df.to_excel('dddddddata1321000-1322000.xlsx')
    #df.to_excel('data1320000-1325000.xlsx',index=False,header=False,startrow=len(reader)+1)
    '''
    data_list = []
    table_row = table.find_all('a', {'class': 'querylabel'})
    print(len(table_row))
    try:
        table_data = table.find_all('td')
        print(len(table_data))
        for m in table_data:
            data_list.append(m.find(('font', {'face': 'arial'})).text)
    except Exception as e:
        print(e)
        table_data = table.find_all('font', {'style': 'font-size:80%'})

    print(data_list)
    #for row, data in zip(table_row, table_data):
        #print(row.text, " :", data.text, '\n')
    #browser.close()
    '''
