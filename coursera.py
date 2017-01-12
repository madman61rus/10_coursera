import requests
import json
from datetime import datetime
from lxml import etree
import random
from bs4 import BeautifulSoup
from openpyxl import Workbook

def get_courses_list():
    courses_list = requests.get('https://www.coursera.org/sitemap~www~courses.xml')
    courses_xml = etree.fromstring(courses_list.content)
    courses_urls =  [child[0].text for child in courses_xml]

    return random.sample(courses_urls,20)


def get_course_info(course_slug):
    result_info = {}
    result_info['url'] = course_slug
    html = requests.get(course_slug)
    soup = BeautifulSoup(html.content,'html.parser')
    if soup:
        result_info['title'] = soup.find('div',{'class':'title'}).contents[0]
        result_info['language'] = soup.find('div',{'class':'language-info'}).contents[1]

        if  soup.find('div',{'class' : 'ratings-text bt3-visible-xs'}):
            result_info['stars'] = soup.find('div',{'class' : 'ratings-text bt3-visible-xs'}).contents[0]
        else:
            result_info['stars'] = ''
        try:
            data_from_script = soup.select('script[type="application/ld+json"]')[0].text
            data_json = json.loads(data_from_script)
            startDate = data_json['hasCourseInstance'][0]['startDate']
            endDate = data_json['hasCourseInstance'][0]['endDate']
            startDt = datetime.strptime(startDate,'%Y-%m-%d')
            endDt = datetime.strptime(endDate, '%Y-%m-%d')
            weeks = (endDt - startDt).days // 7
            result_info['start_date'] = startDate
            result_info['weeks'] = weeks
        except IndexError:
            result_info['start_date'] = None
            result_info['weeks'] = None

    return result_info


def output_courses_info_to_xlsx(filepath,cources):
    wb = Workbook()
    ws = wb.active
    ws.cell(column=1, row=1,value='Title')
    ws.cell(column=2, row=1, value='Language')
    ws.cell(column=3, row=1, value='Start Date')
    ws.cell(column=4, row=1, value='Weeks')
    ws.cell(column=5, row=1, value='Stars')
    ws.cell(column=6, row=1, value='Url')

    for row in range (2,len(cources)+1):
        ws.cell(column=1, row=row, value=cources[row - 1]['title'])
        ws.cell(column=2, row=row, value=cources[row-1]['language'])
        ws.cell(column=3, row=row, value=cources[row-1]['start_date'])
        ws.cell(column=4, row=row, value=cources[row - 1]['weeks'])
        ws.cell(column=5, row=row, value=cources[row - 1]['stars'])
        ws.cell(column=6, row=row, value=cources[row - 1]['url'])
    wb.save(filename=filepath)



if __name__ == '__main__':
    courses_info_dict = []

    courses_list = get_courses_list()
    get_course_info(courses_list[0])

    for course in courses_list:
        courses_info_dict.append(get_course_info(course))

    output_courses_info_to_xlsx('test.xlsx',courses_info_dict)

