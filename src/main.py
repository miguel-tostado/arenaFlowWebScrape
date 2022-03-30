from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import Workbook
import time
import os


###################################################################################

option = webdriver.ChromeOptions()
option.add_argument("--headless")

url = "https://arena.flowrestling.org/"
driver = webdriver.Chrome(
    "C:\\Users\mtost\Downloads\chromedriver_win32\chromedriver", options=option
)
driver.get(url)

element = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.CLASS_NAME, "loader-region"))
)

###################################################################################


def getResultsEvents(page):
    html = page
    line = ""
    numEvents = 0
    eventBlockCount = 0
    isResultsEvent = False
    for char in html:
        line += char
        if not isResultsEvent:
            if char == ">":
                if "events open" in line:
                    line = ""
                    eventBlockCount += 1
                    if eventBlockCount == 2:
                        isResultsEvent = True
        if isResultsEvent:
            if "</li>" in line:
                if "NCAA Division" in line:
                    return line
                else:
                    line = ""
    return line


def getURL(resultBlock):
    eventURL = searchInfo(resultBlock, 'href="/event/', '"')
    if "https" not in eventURL:
        eventURL = "https://arena.flowrestling.org/event/" + eventURL
    return eventURL


def searchInfo(event, start, end):
    line = ""
    isRecording = False
    for char in event:
        line += char
        if start in line:
            line = ""
            isRecording = True
        if isRecording:
            if end in line:
                line = line[:-1]
                break
    return line


def openResultPage(eventURL):
    driver.get(eventURL)
    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CLASS_NAME, "brackets-nav"))
    )
    li = driver.find_element_by_xpath("//div[@class='brackets-nav']/ul/li[1]")
    li.click()
    waitElement = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.NAME, "sort_display"))
    )
    drpSort = Select(driver.find_element_by_xpath("//select[@name='sort_display']"))
    drpSort.select_by_visible_text("By Round")
    waitElement = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CLASS_NAME, "section-list"))
    )
    mainHeaders = driver.find_elements_by_xpath("//div[@class='header']")
    return mainHeaders


def getEventData(eventURL):
    mainHeaders = openResultPage(eventURL)
    roundDict = clickRounds(len(mainHeaders) - 1, 0, eventURL, {})
    return roundDict


def clickRounds(length, count, eventURL, roundDict):
    roundObj = roundDict
    if count == length:
        return roundObj
    else:
        mainHeaders = driver.find_elements_by_xpath("//div[@class='header']")
        # mainHeaders = openResultPage(eventURL)
        # time.sleep(1)
        mainHeaders[count].click()
        time.sleep(4)
        roundName = mainHeaders[count].text
        athleteSpeedBlock = driver.find_elements_by_xpath("//span[@class='seed']")
        athleteNameBlock = driver.find_elements_by_xpath(
            "//span[@class='display-name']"
        )
        athleteSeed = []
        athleteName = []
        athleteGroup = []
        for athlete in athleteSpeedBlock:
            athleteSeed.append(athlete.text)
        for athlete in athleteNameBlock:
            athleteName.append(athlete.text)
        # driver.refresh()
        mainHeaders[count].click()
        time.sleep(1)
        for i in range(len(athleteSeed)):
            athleteGroup.append([athleteSeed[i], athleteName[i]])
        roundObj[roundName] = athleteGroup
        print("Finished ", roundName, " ", count + 1, "/", length)
        result = clickRounds(length, count + 1, eventURL, roundObj)
    return result


def writeToExcel(eventData):
    workbook = Workbook()
    sheet = workbook.active

    count = 1
    data = eventData

    # print(data)
    # print(type(data))

    for event in data:
        sheet["A" + str(count)] = event
        count += 1
        for athlete in data[event]:
            if athlete[0] != "" and athlete[1] != "":
                sheet["A" + str(count)] = athlete[0]
                sheet["B" + str(count)] = athlete[1]
                count += 1

    file = "roundData.xlsx"
    workbook.save(filename=file)
    os.startfile(file)


page = driver.page_source
# driver.quit()

resultBlock = getResultsEvents(page)
resultURL = getURL(resultBlock)
eventData = getEventData(resultURL)
writeToExcel(eventData)


driver.quit()
