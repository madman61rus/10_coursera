import requests
import json
from _datetime import datetime
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


    return result_info


def output_courses_info_to_xlsx(filepath,cources):
    wb = Workbook()
    ws = wb.create_sheet(title="Courses")
    dict = ['title', 'language', 'start_date','weeks','stars','url']
    print(cources[0][dict[0]])
    print(cources[0][dict[1]])
    for row in range (1,20):
        for col in range(1,6):
            ws.cell(column=col, row=row, value=cources[row-1][dict[col-1]])
        print (cources[row-1][dict[col-1]])
    wb.save(filename=filepath)



if __name__ == '__main__':
    courses_info_dict = []

    courses_list = get_courses_list()
    get_course_info(courses_list[0])

    for course in courses_list:
        courses_info_dict.append(get_course_info(course))

    output_courses_info_to_xlsx('test.xslx',courses_info_dict)

