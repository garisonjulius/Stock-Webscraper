import gspread
import time
from selenium import *
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
from decimal import Decimal
import re
import pandas as pd

# Set up Google Sheet
sa = gspread.service_account(filename='/Users/garisonjulius/Downloads/stock/revisedstock-431b2c269022.json')
sh = sa.open_by_url('https://docs.google.com/spreadsheets/d/1v5FbfCuueVbqhKU74Nyd9DKXheI5uXTJ9oIYwX6_-mQ/edit#gid=0')
worksheet = sh.sheet1
worksheet2 = sh.worksheet('Missed')

# Optimizing the performance
options = Options()
#options.add_argument("--headless")  # Run Chrome in headless mode (no GUI)
prefs = {"profile.managed_default_content_settings.images": 2}
options.add_experimental_option("prefs", prefs)
options.add_argument('--disable-extensions')
options.add_argument('start-maximized')
options.add_argument('disable-infobars')
options.add_argument('--disable-popup-blocking')  # Disable pop-ups
options.add_argument('--disable-notifications')   # Disable notifications
options.add_argument('--mute-audio')              # Mute audio
options.add_experimental_option('prefs', {
    'profile.default_content_setting_values.media_stream': 1,  # Disable media stream
    'profile.default_content_setting_values.plugins': 1,       # Disable plugins
    'profile.default_content_setting_values.geolocation': 2,   # Disable geolocation
    'profile.default_content_setting_values.notifications': 2, # Disable notifications
})

# Launch the Chrome browser
PATH = "/Users/garisonjulius/Downloads/stock/chromedriver-mac-arm64/chromedriver"

# Set the ChromeDriver service
service = Service(PATH)

# Create a function for scraping data with Selenium
def scrape_data_with_selenium(driver, url, inpStr, missed):

    try:
        # Zacks page
        driver.get('https://www.zacks.com/stock/quote/' + url + '/detailed-earning-estimates')

        # Find the elements and extract data
        quote = driver.find_elements(By.ID, 'quote_ribbon_v2')
        dataQuote = [info.text for info in quote]
        dataQuote = [val.split('\n') for val in dataQuote]

        twoCol = driver.find_elements(By.ID, 'detailed_estimate')
        dataTwoCol = [data.text for data in twoCol]
        dataTwoCol = [val.split('\n') for val in dataTwoCol]

        estimates = driver.find_elements(By.CLASS_NAME, 'quote_body_full')
        dataEstimates = [data.text for data in estimates][1]

        try:
            dataLst = dataEstimates.split()[52:234]
            indexToKeep = [0, 1, 22, 28, 54, 64, -1]
            dataLst = [dataLst[i] for i in indexToKeep]
        except:
            dataLst = []

        

        finData = dataQuote + dataTwoCol
        finData = [item for sublist in finData for item in sublist]
        delimeter = ' '

        # prefixes_to_find = ["Current Qtr " , "Next Qtr ", "Current Year ", "Next Year ", "PE ", "PEG Ratio"]
        # 39

    
        # Clean Up
        finData[2] = finData[2][:-4]

        if (finData[5].startswith('Add')):
            finData.insert(5, None)

        if finData[9] == 'Style Scores:':
            temp = finData[9:]
            finData = finData[:9] + ['filler'] * 2 + ['NA'] + temp

        if (finData[14][1] == ' '):
           finData[14] = finData[14][33:36]
        else:
            finData[14] = 'NA'

        finData[18] = finData[18][:-16]
        finData[19] = finData[19][10:]
        finData[27] = finData[27][15:]
        finData[28] = finData[28][16:]
        finData[29] = finData[29][17:]
        finData[30] = finData[30][3:]
        finData[34] = finData[34][12:]
        finData[35] = finData[35][9:]  # EPS Next Year

        if not (finData[38].startswith('*BMO')):
            finData.insert(38, None)


        #print(finData)
        #print(finData[40])
        #print('**************************')

        #Growth Estimates
        
        store = finData[40]
        if(len(finData[40].split()) > 6):
            finData[40] = finData[40].split()[5]
        else:
            finData[40] = finData[40].split()[4]

        if(len(finData[41].split()) > 6):
            finData[41] = finData[41].split()[5]
        else:
            finData[41] = finData[41].split()[4]
        
        if(len(finData[42].split()) > 6):
            finData[42] = finData[42].split()[4]
        else:
            finData[42] = finData[42].split()[3]

        if(len(finData[43].split()) > 6):
            finData[43] = finData[43].split()[4]
        else:
            finData[43] = finData[43].split()[3]

        finData[46] = finData[46].split()[1]
        finData[47] = finData[47].split()[2]
        finData = finData[:48]

        finData = finData + dataLst

        # splitting the ticker code and company name
        finData.append(delimeter.join(finData[0].split()[:-1]))
        finData.append(finData[0].split()[-1])
        finData[-1] = finData[-1][1:-1]
        tik = finData[-1]

        print('FINISHED ZACKS')

        finData += ["FILLER"] * 9
        finData[11] = finData[11].strip()


        if(len(inpStr) > 0):
            finData.append(inpStr)
        else:
            finData.append("N/A")

        print('FINISHED YAHOO')

        #Wallmine
        innerFinData = []
    
        driver.get('https://wallmine.com')
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="search-form"]/div/span/input[2]')))
        searchBar = driver.find_element(By.XPATH, '//*[@id="search-form"]/div/span/input[2]')
        searchBar.send_keys(tik)
        searchBar.submit()

        print('CHECK')

        #Making sure we get the right company data from wallmine

        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/main/section/div[4]')))
        print(1)
        check = driver.find_elements(By.XPATH, '/html/body/main/section')
        print(2)
        checkStats = [info.text for info in check]
        print(3)
        checkStats = [val.split('\n') for val in checkStats][0]
        
        #print(checkStats)

        checkTik = checkStats[0].split()[0]

        print(checkTik)
        print(tik)

        if (str(checkTik).lower() != str(tik).lower()):
            print('MISMATCH')
            return finData

        print('CHECK WEEEE')

        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/main/section/div[4]')))
        value = driver.find_elements(By.XPATH, '/html/body/main/section/div[4]/div[1]/div[2]')
        valueStats = [info.text for info in value]
        valueStats = [val.split('\n') for val in valueStats][0]

        #print(valueStats)

        if (len(valueStats) < 2):
            valueStats = getSecondTryData(driver, url, inpStr, missed)
            #valueStats = []
        else:
            valueStats = valueStats[16:18] + valueStats[24:26] + valueStats[30:32] + valueStats[52:58] + valueStats[64:68] + valueStats[70:72] + valueStats[74:76] + valueStats[88:92] + valueStats[118:120] + valueStats[82:84] + valueStats[12:14] + valueStats[-3:-1] + valueStats[10:12]

        #Revenue Q/Q = 10-11
        #Revenue Y/Y = 12-13
        #Forward/value = 16-17
        #PEG = 24-25
        #P/B = 30-31
        #FCF yeild = 36-37
        #EPS Q/Q = 48-49
        #3 Margins = 52-57
        #ROE + ROI = 64-67
        #D/E = 70-71
        #C Ratio = 74-75
        #RSI = 82-83
        #SMA = 88-91
        #Earning Date = 117-118
        #Country = -3 - -2

        #32 length
        #print(len(valueStats))
        
        #If the second try of getting wallmine data fails, then just return zacks
        if (len(valueStats) < 2):
            return finData

        valueTwo = driver.find_elements(By.XPATH, '/html/body/main/section/div[4]/div[2]/div[1]')
        valueTwoStats = [info.text for info in valueTwo]
        valueTwoStats = [val.split('\n') for val in valueTwoStats][0]

        
        valueTwoStats = valueTwoStats[6:12]
        
        #print(valueTwoStats)
     
        #icon-sign wmi wmi-caret-down
    
        print('****************')
        
        finData = finData + valueTwoStats + valueStats

        threeMonthNegative = driver.find_element(By.XPATH, '/html/body/main/section/div[4]/div[2]/div[1]/div[4]/a/div/div/i')
        threeClassName = threeMonthNegative.get_attribute('class')

        if(threeClassName == 'icon-sign wmi wmi-caret-down'):
            #print(finData[68]) #28.92%
            finData[68] = Decimal(finData[68][:-1])
            finData[68] = finData[68] * -1
            finData[68] = str(finData[68]) + '%'

        oneYearNegative = driver.find_element(By.XPATH, '/html/body/main/section/div[4]/div[2]/div[1]/div[6]/a/div/div/i')
        oneClassName = oneYearNegative.get_attribute('class')
        
        if(oneClassName == 'icon-sign wmi wmi-caret-down'):
            #print(finData[70]) #4.14%
            finData[72] = Decimal(finData[70][:-1])
            finData[72] = finData[70] * -1
            finData[72] = str(finData[70]) + '%'

        print("FINISHED WALLMINE")

        #Final edits

        if(len(store.split()) > 6):
            store = store.split()[4]
        else:
            store = store.split()[3]

        finData.append(store)

        #Making sure wrong company data is not gathered
        finData.append(checkTik)

        print(finData)
        print('\n')
        print(len(finData))

        return finData

    except Exception as e:
        missed.append(url)
        print("Error occurred: lol", e)
        return finData

def getSecondTryData(driver, url, inpStr, missed):
    
    #Second attempt at getting wallmine data if the first time fails

    print('---------------')
    print("REPEAT REPEAT REPEAT")
    print('---------------')

    try:

        print('FINISHED YAHOO')

        #Wallmine
        innerFinData = []
        driver.get('https://wallmine.com')
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="search-form"]/div/span/input[2]')))
        searchBar = driver.find_element(By.XPATH, '//*[@id="search-form"]/div/span/input[2]')
        searchBar.send_keys(url)
        searchBar.submit()

        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/main/section/div[4]')))
        value = driver.find_elements(By.XPATH, '/html/body/main/section/div[4]/div[1]/div[2]')
        valueStats = [info.text for info in value]
        valueStats = [val.split('\n') for val in valueStats][0]

        if (len(valueStats) < 2):
            #innerFinData = getSecondTryData(driver, url, inpStr, missed)
            return []
            #valueStats = []
        
        else:
            valueStats = valueStats[16:18] + valueStats[24:26] + valueStats[30:32] + valueStats[52:58] + valueStats[64:68] + valueStats[70:72] + valueStats[74:76] + valueStats[88:92] + valueStats[118:120] + valueStats[82:84] + valueStats[12:14] + valueStats[-3:-1] + valueStats[10:12]
            return valueStats
    
        #Revenue Y/Y = 12-13
        #Forward/value = 16-17
        #PEG = 24-25
        #P/B = 30-31
        #FCF yeild = 36-37
        #EPS Q/Q = 48-49
        #3 Margins = 52-57
        #ROE + ROI = 64-67
        #D/E = 70-71
        #C Ratio = 74-75
        #RSI = 82-83
        #SMA = 88-91
        #Earning Date = 117-118
        #Country = -3 - -2

        valueTwo = driver.find_elements(By.XPATH, '/html/body/main/section/div[4]/div[2]/div[1]')
        valueTwoStats = [info.text for info in valueTwo]
        valueTwoStats = [val.split('\n') for val in valueTwoStats][0]
        valueTwoStats = valueTwoStats[6:8] + valueTwoStats[10:12]
        
        #icon-sign wmi wmi-caret-down
        
        if(len(innerFinData) < 2):
            finData = finData + valueTwoStats + valueStats
        else:
            finData = innerFinData

        print("FINISHED WALLMINE REPEAT")
        #missed.remove(url)
        return finData

    except Exception as e:
        #missed.append(url)
        print("Error occurred: lmao", e)
        return None


# Create a function for processing specific stocks
def process_specific_stocks(driver, missed):
    tik_lst = []

    while True:
        user_input = input("Enter Stock Codes (type 'quit' to exit): ")

        if user_input.lower() == 'quit':
            break

        tik_lst.append(user_input.upper().strip())

    print(tik_lst)
    print("Calculating Now...")

    for code in tik_lst:
        
        try:
            basic_info = scrape_data_with_selenium(driver, code, "", missed)

            print('ONE BY ONE')
            print(basic_info)

            time.sleep(2)
            if basic_info:
                print(basic_info)
                print(len(basic_info))
                #print(basic_info.index('3 months'))

                if len(basic_info) < 60:
                    print('BAD DATA')
                    continue
                elif(len(basic_info) < 102):
                    worksheet.insert_row(basic_info[:67], 3)
                else:
                    worksheet.insert_row(basic_info, 3)
            else:
                print("Failed to scrape data from the given URL.")
        except TimeoutException:
            print(f"Timeout occurred while processing link: {code}. Moving on to the next link.")
            continue
        except Exception as e:
            print(f"Error occurred while processing link: {code}\nError: {e}")
            continue

# Create a function for processing the calendar
def process_calendar(driver, missed):
    inp = input('What Day: ')
    inp = int(inp)
    inpStr = str(inp)

    # Open the website
    url = "https://www.zacks.com/earnings/earnings-calendar?icid=home-home-nav_tracking-zcom-main_menu_wrapper-earnings_calendar"
    driver.get(url)

    #time.sleep(1)

    # Go to the specific date using explicit wait
    link = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="date_select"]')))
    driver.execute_script("arguments[0].click();", link)

    time.sleep(5)
    

    dateLink = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="dt_' + inpStr + '"]')))
    driver.execute_script("arguments[0].click();", dateLink)

    time.sleep(5)
 

    # View all the symbols using Select class
    allLink = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="earnings_rel_data_all_table_length"]/label/select')))
    all_select = Select(allLink)
    all_select.select_by_value("-1")

    # Get the total number of entries
    time.sleep(2)
    
    entry_info = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="earnings_rel_data_all_table_info"]')))
    entries_text = entry_info.text.split()[-2]
    entries = int(entries_text)
    print("Number of entries  " + str(entries))

    # Wait for the links to be present (optional, if the page is dynamically loaded)
    try:
        WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.TAG_NAME, 'a')))
    except:
        print("Links not found within the given time")

    # Scrape all the symbols using find_elements_by_...
    links = driver.find_elements(By.TAG_NAME, 'a')

    time.sleep(3)
    
    filtered_links = [link.get_attribute('href') for link in links if
                      link.get_attribute('href') and 'stock/quote' in link.get_attribute('href')]
    filtered_links = filtered_links[:entries]

    #print(filtered_links)

    #time.sleep(2)
    print("Number of filtered links  " + str(len(filtered_links)))
    # filtered_links = ['https://www.zacks.com/stock/quote/AAPL', 'https://www.zacks.com/stock/quote/PLTR', 'https://www.zacks.com/stock/quote/MSFT']
    # Call the function for each link and accumulate data in rows_to_insert
    for value in filtered_links:
        try:
            basic_info = scrape_data_with_selenium(driver, value[34:], inpStr, missed)

            if basic_info:
                #print(basic_info)
                if(len(basic_info) < 102):
                    worksheet.insert_row(basic_info[:67], 3)
                else:
                    worksheet.insert_row(basic_info, 3)
            else:
                print("Failed to scrape data from the given URL.")
        except TimeoutException:
            print(f"Timeout occurred while processing link: {value}. Moving on to the next link.")
            continue
        except Exception as e:
            print(f"Error occurred while processing link CALENDAR: {value}\nError: {e}")
            continue
    

# Main script
#userInput = input('1 for Specific Stocks - C for Calendar:  ').strip()
print('HERE')
missed = []


with webdriver.Chrome(service=service, options=options) as driver:

    driver.get('https://finviz.com/quote.ashx?t=MSFT&p=d')
    finVizPath = '/html/body/div[4]/div[3]/div[4]/table/tbody/tr/td/div/table[1]/tbody/tr/td/div[2]/table/tbody/tr'

    #13 rows
    finVizFinal = []
    nested = []

    for i in range(1, 14):
        nums = driver.find_element(By.XPATH, finVizPath + '[' + str(i) + ']')
        text = str(nums.text).split(' ')
        print(text)
        nested.append(text)

    '''#First Row
    first = nested[0]
    peIndex = first.index('P/E')
    finVizFinal.append(first[peIndex + 1])'''

    #Second Row 
    second = nested[1]
    forwardIndex = second.index('P/E')
    finVizFinal.append(second[forwardIndex + 1])

    #Third Row
    third = nested[2]
    pegIndex = third.index('PEG')
    finVizFinal.append(third[pegIndex + 1])

    quarterlyIndex = third.index('Quarter')
    finVizFinal.append(third[quarterlyIndex + 1])

    #Fourth Row
    fourth = nested[3]
    halfIndex = fourth.index('Half')
    finVizFinal.append(fourth[halfIndex + 2])

    #Fifth Row
    fifth = nested[4]
    priceBookIndex = fifth.index('P/B')
    finVizFinal.append(fifth[priceBookIndex + 1])

    annualIndex = fifth.index('Year')
    finVizFinal.append(fifth[annualIndex + 1])

    #Sixth Row
    sixth = nested[5]
    roeIndex = sixth.index('ROE')
    finVizFinal.append(sixth[roeIndex + 1])

    #Seventh Row
    seventh = nested[6]
    roiIndex = seventh.index('ROI')
    finVizFinal.append(seventh[roiIndex + 2])

    #Eigth Row
    eigth = nested[7]
    grossIndex = eigth.index('Gross')
    finVizFinal.append(eigth[grossIndex + 2])

    #Ninth Row
    ninth = nested[8]
    currentIndex = ninth.index('Current')
    finVizFinal.append(ninth[currentIndex + 2])

    operatingIndex = ninth.index('Oper.')
    finVizFinal.append(ninth[operatingIndex + 2])

    rsiIndex = ninth.index('RSI')
    finVizFinal.append(ninth[rsiIndex + 2])

    #Tenth Row
    tenth = nested[9]
    debtEquityIndex = tenth.index('Debt/Eq')
    finVizFinal.append(tenth[debtEquityIndex + 1])

    yearOverYearIndex = tenth.index('Y/Y')
    finVizFinal.append(tenth[yearOverYearIndex + 2])

    netIndex = tenth.index('Profit')
    finVizFinal.append(tenth[netIndex + 2])

    #Twelfth Row
    twelfth = nested[11]
    quarterIndex = twelfth.index('Q/Q')
    finVizFinal.append(twelfth[quarterIndex + 1])

    #Thirteenth Row
    thirteenth = nested[12]
    fiftyIndex = thirteenth.index('SMA50')
    finVizFinal.append(thirteenth[fiftyIndex + 1])

    twoHundredIndex = thirteenth.index('SMA200')
    finVizFinal.append(thirteenth[twoHundredIndex + 1])

    #Print Final Data
    print(finVizFinal)
    #worksheet2.insert_row(finVizFinal, 3)

    


    

#worksheet2.insert_row(missed, 3)
print("Data inserted into Google Sheets.")