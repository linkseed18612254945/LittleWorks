import re
from bs4 import BeautifulSoup
import requests
import csv
import time

COL_NAME = ['patent_code', 'patent_name', 'year', 'inventor_and_country_data', 'Assignee', 'abstract', 'IPC', 'USPC', 'CPC']

# 获取专利数据
def get_patent_data(url):
    patents =[]
    tmp_s = requests.session()
    try:
        r2 = tmp_s.get(url)
    except:
        print('There is something wrong with this patent')
        return
    text2 = r2.text
    tmp_soup = BeautifulSoup(text2, "html.parser")
    patent_data = {}
    patent_data['patent_code'] = tmp_soup.find('title').next[22:]# 找到了title标签，取22位及以后的data
    patent_data['patent_name'] = tmp_soup.find('font', size="+1").text[:-1].replace('\n ','').replace('    ',' ')# 找到font标签，size=“+1”，获取全部文本
    tmp1 = text2[re.search('BUF7=', text2).span()[1]:]
    patent_data['year'] = tmp1[:re.search('\n', tmp1).span()[0]]
    patent_data['inventor_and_country_data'] = tmp_soup.find_all('table', width="100%")[2].contents[1].text.replace('Inventors: ','').strip()
    patent_data['abstract'] = str(tmp_soup.find('p').get_text()).replace('\n ','').replace('    ',' ')
    # patent_data['USPC'] = tmp_soup.find_all('table', width="100%")[5].contents[1].text.replace("Current U.S. Class: ","")
    # patent_data['CPC'] = tmp_soup.find_all('table', width="100%")[5].contents[3].get_text().replace("Current CPC Class: ","")
    # patent_data['IPC'] = tmp_soup.find_all('table', width="100%")[5].contents[5].get_text().replace("Current International Class: ","")
    tabletags_list =tmp_soup.findAll("table")
    ## US CLASSES -- could be in 4th,5th, 6th or sometimes 7th table
    clss_tbl = 4
    # Find out if we have the right table index
    if (tabletags_list[clss_tbl].findAll(text="Current U.S. Class:") == []):
        clss_tbl = 5
    if (tabletags_list[clss_tbl].findAll(text="Current U.S. Class:") == []):
        clss_tbl = 6
    if (tabletags_list[clss_tbl].findAll(text="Current U.S. Class:") == []):
        clss_tbl = 7
    usclass_str = "".join(tabletags_list[clss_tbl].extract().findAll('td')[1].findAll(text=True))
    patent_data['USPC'] = usclass_str
    CPC_str = "".join(tabletags_list[clss_tbl].extract().findAll('td')[3].findAll(text=True))
    patent_data['CPC'] = CPC_str
    IPC_str="".join(tabletags_list[clss_tbl].extract().findAll('td')[5].findAll(text=True))
    patent_data['IPC']=IPC_str
    assi = 2
    if (tabletags_list[assi].findAll(text="Assignee:") == []):
        assi = 3
    if (tabletags_list[assi].findAll(text="Assignee:") == []):
        assi = 4
    if (tabletags_list[assi].findAll(text="Assignee:") == []):
        assi = 5
    Assignee = tmp_soup.find(text="Assignee:")
    try:
        patent_data['Assignee'] = Assignee.find_parent().find_parent().td.get_text()
    except:
        patent_data['Assignee'] = ''
    patent_data_list = [patent_data['patent_code'], patent_data['patent_name'],patent_data['year'], patent_data['inventor_and_country_data'],patent_data['Assignee'],patent_data['abstract'],patent_data['IPC'],patent_data['USPC'],patent_data['CPC']]
    return patent_data_list


def create_dict(patent_list):
    patent_dict = {}
    for i in range(9):
        patent_dict[COL_NAME[i]] = patent_list[i]
    return patent_dict


def store_patent_data(patents):
    date_time = time.strftime("%Y%m%d_%H%M%S", time.localtime(time.time()))
    file_name = 'PatentData' + '_' + str(date_time) +'.csv'
    with open(file_name, 'w', encoding='utf-8') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=COL_NAME, restval='NULL')
        writer.writeheader()
        writer.writerows(patents)


def main():
    patents = []
    patent_num = [4001]
    count = 0
    start_time = time.time()
    try:
        for i in patent_num:
            html = "http://patft.uspto.gov/netacgi/nph-Parser?Sect1=PTO1&Sect2=HITOFF&d=PALL&p=1&u=%2Fnetahtml%2FPTO%2Fsrchnum.htm&r=1&f=G&l=50&s1=8230123.PN.&OS=PN/8230123&RS=PN/8230123"
            patent_data_list = get_patent_data(html)
            now_time = time.time()
            cost_time = now_time - start_time
            print('Processing the number ' + str(i) + ' patent. ', 'Cost time : %.2f s' % cost_time)
            patent_dict = create_dict(patent_data_list)
            patents.append(patent_dict)
            count += 1
    except:
        print('ERROR! in the number ' + str(count+1) + ' patent')
    store_patent_data(patents)
    now_time = time.time()
    cost_time = now_time - start_time
    print('Processing Complete. Processed ' + str(count) + ' patents', 'Cost time : %.2f s' % cost_time)

if __name__ == '__main__':
    main()

