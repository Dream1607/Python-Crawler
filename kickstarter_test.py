import urllib2
import re
import math
import datetime
import xlrd
import xlwt
from xlutils.copy import copy
from bs4 import BeautifulSoup

#######open excel
# workbook = xlrd.open_workbook('projects.xls')
# worksheet = workbook.sheet_by_name('Sheet1')
# last_row = worksheet.nrows - 1
# last_name = worksheet.row(last_row)[0].value

class Proj:
    def __init__(self,name,link,crawl_time):
        self.name = name
        self.link = link
        self.crawl_time = crawl_time

pageNum = 1
flag = True
newestProjects = []

while flag:
    url = "https://www.kickstarter.com/discover/advanced?google_chrome_workaround&page="+str(pageNum)+"&category_id=0&woe_id=0&sort=newest"
    pageNum += 1
    html = urllib2.urlopen(url).read()
    soup = BeautifulSoup(''.join(html))

    projects = soup.findAll("h6", { "class": "project-title mobile-center" })

    for project in projects:
        if project.findNext('a').string == last_name:
            flag = False
        if flag:
            newestProjects.append(Proj(project.findNext('a').string, "http://www.kickstarter.com" + project.findNext('a')['href'], datetime.datetime.now()))

######output to excel
# workbook = copy(workbook)
# worksheet = workbook.get_sheet(0)
#
# for i in range(len(newestProjects)):
#     worksheet.write(i+last_row+1, 0, newestProjects[len(newestProjects)-i-1].name)
#     worksheet.write(i+last_row+1, 1, newestProjects[len(newestProjects)-i-1].link)
#     worksheet.write(i+last_row+1, 2, newestProjects[len(newestProjects)-i-1].crawl_time)
# workbook.save("projects.xls")
