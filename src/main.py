from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import Workbook
import time
import os


###################################################################################
####    Initiates the web driver for Chrome

option = webdriver.ChromeOptions()
option.add_argument("--headless")
option.add_argument("--log-level=3")

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
    """Accepts the HTML from the original URL as a string and goes through each character.
    As it scans, each character is added to the "line" variable.
    Once a closing tag ">" is found, the "line" is scanned for the "events open" CSS class-name.
    If found, it will increase the numEvents counter by 1 (there are two sections that contain this class name: Upcoming and Results).
    If "events open" is not found, then the line is deleted and the next HTML tag is read.
    Once the second "events open" CSS class is found, isResultsEvent is then "True" which will trigger the next section of code:
    it will continue to scan as before, but now it's looking for the closing "</li>" tag. If "NCAA Division" is found in that list item,
    then it will return the line.

    Note: The website scanned rarely uses ID's so I had to search by class-name.

    Args:
        page (string): The source HTML as a string

    Returns:
        string: Returns everything within the <li> </li> tags that contain the word "NCAA Division"
    """
    line = ""
    numEvents = 0
    isResultsEvent = False
    for char in page:
        line += char
        if not isResultsEvent:
            if char == ">":
                if "events open" in line:
                    line = ""
                    numEvents += 1
                    if numEvents == 2:
                        isResultsEvent = True
        if isResultsEvent:
            if "</li>" in line:
                if "NCAA Division" in line:
                    return line
                else:
                    line = ""
    return line


def getURL(resultBlock):
    """Accepts the returned string from the getResultsEvents function and sends it to the searchInfo function which returns the specific
    "/event/..." portion of the URL.
    This function will concatonate the returned port of the URL with the base URL to create the specific event's full URL.

    Args:
        resultBlock (string): Returned string from getResultsEvents which contains everything within the <li> and </li> tags for this event

    Returns:
        string: Returns the specific event's full URL
    """
    eventURL = searchInfo(resultBlock, 'href="/event/', '"')
    if "https" not in eventURL:
        eventURL = "https://arena.flowrestling.org/event/" + eventURL
    return eventURL


def searchInfo(event, start, end):
    """Takes the string from getURL and the search criteria to locate the event's URL.
    This function will scan each character in the string and add it to the line variable.
    As each character is added, the function will check if the "start" arg is has been scanned. If it has, the function will turn on recording (isRecording) and delete the string inside line. While recording, the desired information will stored in line until the string specified in "end" has been scanned.
    Once completed, the function will return the desired information (event's URL)

    Args:
        event (string): String of the HTML information: everything between the <li> </li> tags.
        start (string): What the scanner will look for to start recording the string.
        end (string): What the scanner will look for to end recording the string.

    Returns:
        string: Returns the desired URL segment
    """
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


def getEventData(eventURL):
    """First call to scan the Results page. This is mainly just to setup the clickRounds function.

    Args:
        eventURL (string): Event URL

    Returns:
        dictionary: Returns the final dictionary object that is returned by clickRounds
    """
    mainHeaders = openResultPage(eventURL)
    roundData = clickRounds(len(mainHeaders) - 1, 0, eventURL, {})
    return roundData


def openResultPage(eventURL):
    """Unfortunately, arena.flowrestling.org pull all of its information dynamically; which means that I, with my current skill level, have to refresh the page after scanning each round's data. I initially tried to simply simulate closing that rounds' section, search for the "header" class-name again to refresh the list, and clicking on the next round. That does work; however, it kept missing the very first match in a few rounds. By refreshing the page, searching for the "header" class-name again, and clicking the next round, I'm able to get everything.

    This function will cause the page to refresh when called, go through the process of getting back to the "Results" tab, and refresh the element list for tags that have the "header" class.

    Args:
        eventURL (string): The event's URL

    Returns:
        selenium web element: Returns the mainHeaders element array
    """
    driver.get(eventURL)
    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CLASS_NAME, "brackets-nav"))
    )
    li = driver.find_element(By.XPATH, "//div[@class='brackets-nav']/ul/li[1]")
    li.click()
    waitElement = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.NAME, "sort_display"))
    )
    drpSort = Select(driver.find_element(By.XPATH, "//select[@name='sort_display']"))
    drpSort.select_by_visible_text("By Round")
    waitElement = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CLASS_NAME, "section-list"))
    )
    mainHeaders = driver.find_elements(By.XPATH, "//div[@class='header']")
    return mainHeaders


def clickRounds(length, count, eventURL, roundDict):
    """Recursively function that will click on the next round after each page refresh. This function will pull all the information from each round and store it in the "results" dictionary to then be written to an Excel spreadsheet.

    Args:
        length (int): Length of the mainHeaders array. Used to determine how many headers there are.
        count (int): Keeps track of which header is being checked. Compared with length to determin when done scanning.
        eventURL (str): Event's URL. Used to re-open the page after each refresh.
        roundDict (dict): Contains all the data from each round.

        {roundName: [[athleteSeed, athleteName], [athleteSeed, athleteName], ...],
        roundName: ...}

    Returns:
        dict: Returns the result dictionary.
    """
    if count == length:
        return roundDict
    else:
        mainHeaders = openResultPage(eventURL)
        time.sleep(1)
        mainHeaders[count].click()
        time.sleep(3)

        roundName = mainHeaders[count].text
        athleteSeedBlock = driver.find_elements(By.XPATH, "//span[@class='seed']")
        athleteNameBlock = driver.find_elements(
            By.XPATH, "//span[@class='display-name']"
        )
        athleteSeed = []
        athleteName = []
        roundAthleteInfo = []

        for athlete in athleteSeedBlock:
            athleteSeed.append(athlete.text)
        for athlete in athleteNameBlock:
            athleteName.append(athlete.text)

        driver.refresh()
        time.sleep(1)

        for i in range(len(athleteSeed)):
            roundAthleteInfo.append([athleteSeed[i], athleteName[i]])

        roundDict[roundName] = roundAthleteInfo

        print("Finished ", roundName, " ", count + 1, "/", length)
        result = clickRounds(length, count + 1, eventURL, roundDict)
    return result


def writeToExcel(eventData):
    """Writes the data from clickRounds into an Excel spreadsheet.

    Args:
        eventData (dict): Event dictionary that uses each round as a key, and each athlete has it's own array inside the event array.
    """
    workbook = Workbook()
    # sheet = workbook[""]
    sheet = workbook.active

    count = 1
    data = eventData

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

resultBlock = getResultsEvents(page)
resultURL = getURL(resultBlock)
eventData = getEventData(resultURL)
writeToExcel(eventData)


driver.quit()
