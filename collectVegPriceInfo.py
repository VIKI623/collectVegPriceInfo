#coding=utf-8
import requests
import math
import xlwt
from bs4 import BeautifulSoup


url_base = "http://www.xinfadi.com.cn"
prereadydata = [u"蔬菜",u"水果",u"肉禽蛋",u"水产",u"粮油"]

def getKindSet():
    try:
        print "Note:1-蔬菜,2-水果,3-肉禽蛋,4-水产,5-粮油\n"
        print "Example:If you want to search 蔬菜, 水果 and 肉禽蛋, you should input 1&2&3\n"
        kindSet = raw_input("Please input the kind set you want to search:\n")
        return kindSet
    except:
        return "Error"

def getTimeRange():
    try:
        print "Please input the time range you want to search, like 2018-03-18\n"
        print "Note:beginTime is included and endTime isn't included\n"
        beginTime = raw_input("Please input beginTime:\n")
        endTime = raw_input("Please input endTime:\n")
        timeRange = "begintime=" + str(beginTime) + "&endtime=" + str(endTime)
        return timeRange
    except:
        return "Error"

def getHtml(url):
    try:
        r = requests.get(url)
        r.raise_for_status()
        r.encoding = r.apparent_encoding
        return r.text
    except:
        return "Error"

if __name__ == "__main__":
    kindSet = getKindSet()
    if kindSet == "Error":
        print "Failed to get kind set!\n"
    else:
        timeRange = getTimeRange()
        if timeRange == "Error":
            print "Failed to get time range!\n"
        else:
            x = 0
            y = 0
            wb = xlwt.Workbook()
            fileName = timeRange + ".xls"
            kindArray = kindSet.split('&')
            for kind in kindArray:
                pageNum = 1
                url = url_base + "/marketanalysis/" + kind + "/list/" + str(pageNum) + ".shtml?prodname=&" + timeRange
                html = getHtml(url)
                if html == "Error":
                    print "Failed to open web page!\n"
                else:
                    soup = BeautifulSoup(html, "html.parser")
                    content = soup.find('div', class_="hangq_left")
                    resultNum = int(content.em.em.get_text())
                    if resultNum == 0:
                        print "There is not any results!\n"
                    else:
                        ws = wb.add_sheet(prereadydata[int(kind) - 1])
                        pageTotal = math.ceil(resultNum / 20.0)

                        print "Page " + str(pageNum) + " is loading......\n"
                        tableData = content.find_all('tr')
                        for data in tableData:
                            rowData = data.find_all('td')
                            for j in range(0, len(rowData) - 1):
                                ws.write(x, y, rowData[j].get_text())
                                y = y + 1
                            y = 0
                            x = x + 1
                        print "Page " + str(pageNum) + " has been collected!\n"
                        pageNum = pageNum + 1

                        while pageNum <= pageTotal:
                            print "Page " + str(pageNum) + " is loading......\n"
                            url = url_base + "/marketanalysis/" + kind + "/list/" + str(pageNum) + ".shtml?prodname=&" + timeRange
                            html = getHtml(url)
                            if html == "Error":
                                print "Failed to open web page!\n"
                                wb.save(fileName)
                                break
                            else:
                                soup = BeautifulSoup(html, "html.parser")
                                content = soup.find('table', class_="hq_table")
                                tableData = content.find_all('tr')
                                for i in range(1, len(tableData)):
                                    rowData = tableData[i].find_all('td')
                                    for j in range(0, len(rowData) - 1):
                                        ws.write(x, y, rowData[j].get_text())
                                        y = y + 1
                                    y = 0
                                    x = x + 1
                            print "Page " + str(pageNum) +" has been collected!\n"
                            pageNum = pageNum + 1
                        x = 0
                        print "\nThe kind " + prereadydata[int(kind) - 1] + " has been collected!\n"
                        print "#############################################################"
        wb.save(fileName)
        print "\n\nAll data is saved successfully as a file named " + str(timeRange) + "!!!\n"
        raw_input("Please press any key to exit......\n")