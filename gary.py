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
        time.sleep(1)

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

        finData += ["FILLER"] * 9
        finData[11] = finData[11].strip()


        if(len(inpStr) > 0):
            finData.append(inpStr)
        else:
            finData.append("N/A")

        #Final edits
        if(len(store.split()) > 6):
            store = store.split()[4]
        else:
            store = store.split()[3]

        finData.append(store)

        print('FINISHED ZACKS')


        #FINVIZ

        driver.get('https://finviz.com/quote.ashx?t=' + url + '&p=d')
        time.sleep(1)

        #13 rows of data 
        nested = []

        #finVizPath = '/html/body/div[4]/div[3]/div[4]/table/tbody/tr/td/div/table[1]/tbody/tr/td/div[2]/table/tbody/tr'
        #finVizPath = '/html/body/div[3]/div[2]/div[4]/table/tbody/tr/td/div/table[1]/tbody/tr/td/div[2]/table/tbody'

        # Loop through each row of the table and extract text
        row = driver.find_elements(By.CLASS_NAME, 'table-dark-row')
        for val in row:
            nested.append(val.text)

        '''#First Row
        first = nested[0]
        peIndex = first.index('P/E')
        finData.append(first[peIndex + 1])'''

        #Second Row 
        second = nested[1]
        second = second.split(' ')
        #print(second)
        forwardIndex = second.index('P/E')
        finData.append(second[forwardIndex + 1])

        #Third Row
        third = nested[2]
        third = third.split(' ')
        #print(third)
        pegIndex = third.index('PEG')
        finData.append(third[pegIndex + 1])

        quarterlyIndex = third.index('Quarter')
        finData.append(third[quarterlyIndex + 1])

        #Fourth Row
        fourth = nested[3]
        fourth = fourth.split(' ')
        #print(fourth)
        halfIndex = fourth.index('Half')
        finData.append(fourth[halfIndex + 2])

        #Fifth Row
        fifth = nested[4]
        fifth = fifth.split(' ')
        #print(fifth)
        priceBookIndex = fifth.index('P/B')
        finData.append(fifth[priceBookIndex + 1])

        annualIndex = fifth.index('Year')
        finData.append(fifth[annualIndex + 1])

        #Sixth Row
        sixth = nested[5]
        sixth = sixth.split(' ')    
        #print(sixth)
        roeIndex = sixth.index('ROE')
        finData.append(sixth[roeIndex + 1])

        #Seventh Row
        seventh = nested[6]
        seventh = seventh.split(' ')
        #print(seventh)
        roiIndex = seventh.index('ROI')
        finData.append(seventh[roiIndex + 1])

        #eight Row
        eight = nested[7]
        eight = eight.split(' ')
        #print(eight)
        grossIndex = eight.index('Gross')
        finData.append(eight[grossIndex + 2])

        #Ninth Row
        ninth = nested[8]
        ninth = ninth.split(' ')
        #print(ninth)
        currentIndex = ninth.index('Current')
        finData.append(ninth[currentIndex + 2])

        operatingIndex = ninth.index('Oper.')
        finData.append(ninth[operatingIndex + 2])

        rsiIndex = ninth.index('RSI')
        finData.append(ninth[rsiIndex + 2])

        #Tenth Row
        tenth = nested[9]
        tenth = tenth.split(' ')
        #print(tenth)
        debtEquityIndex = tenth.index('Debt/Eq')
        finData.append(tenth[debtEquityIndex + 1])

        yearOverYearIndex = tenth.index('Y/Y')
        finData.append(tenth[yearOverYearIndex + 2])

        netIndex = tenth.index('Profit')
        finData.append(tenth[netIndex + 2])

        #Twelfth Row
        twelfth = nested[11]
        twelfth = twelfth.split(' ')
        #print(twelfth)
        quarterIndex = twelfth.index('Q/Q')
        finData.append(twelfth[quarterIndex + 1])

        earningsIndex = twelfth.index('Earnings')
        uno = twelfth[earningsIndex + 1]
        dos = twelfth[earningsIndex + 2]
        tres = twelfth[earningsIndex + 3]
        total = uno + '-' + dos + '-' + tres
        finData.append(total)

        #Thirteenth Row
        thirteenth = nested[12]
        thirteenth = thirteenth.split(' ')
        #print(thirteenth)
        fiftyIndex = thirteenth.index('SMA50')
        finData.append(thirteenth[fiftyIndex + 1])

        twoHundredIndex = thirteenth.index('SMA200')
        finData.append(thirteenth[twoHundredIndex + 1])

        print('FINISHED FINVIZ')

        # #Etrade 

        # driver.get('https://us.etrade.com/etx/mkt/quotes?symbol=' + url + '#/snapshot')   

        # sentimentVal = driver.find_element(By.XPATH , '/html/body/div[2]/div/div[2]/main/div[4]/div/div/div[1]/div[3]/div/div/div[2]/div/div/div/div[2]/div[1]/div[1]/div[2]')
        # print(sentimentVal.text)

        # print('-----')

        # bigger = driver.find_element(By.XPATH , '/html/body/div[2]/div/div[2]/main/div[4]/div/div/div[1]/div[3]/div/div/div[2]/div/div/div/div[2]')
        # print(bigger.text)

        # print('FINISHED ETRADE')

        return finData

    except Exception as e:
        missed.append(url)
        print("Error occurred: lol", e)
        return finData



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
                elif(len(basic_info) < 80):
                    #There should be 86 items in the list
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

    time.sleep(1.5)

    dateLink = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="dt_' + inpStr + '"]')))
    driver.execute_script("arguments[0].click();", dateLink)

    time.sleep(1.5)
 

    # View all the symbols using Select class
    allLink = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="earnings_rel_data_all_table_length"]/label/select')))
    all_select = Select(allLink)
    all_select.select_by_value("-1")

    # Get the total number of entries
    time.sleep(1)
    
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

    time.sleep(1.5)
    
    filtered_links = [link.get_attribute('href') for link in links if
                      link.get_attribute('href') and 'stock/quote' in link.get_attribute('href')]
    filtered_links = filtered_links[:entries]

    #print(filtered_links)

    print("Number of filtered links  " + str(len(filtered_links)))
    # filtered_links = ['https://www.zacks.com/stock/quote/AAPL', 'https://www.zacks.com/stock/quote/PLTR', 'https://www.zacks.com/stock/quote/MSFT']

    #This makes the filtered links look like filtered codes. Like this [AAPL, PLTR, MSFT]
    for i in range (len(filtered_links)): 
        filtered_links[i] = filtered_links[i][34:]
        #print(filtered_links[i])
            
    
    # #STARTING FROM A SPECIFIC STOCK
    # index = filtered_links.index('TSSI')
    # filtered_links = filtered_links[index:]
    # #STARTING FROM A SPECIFIC STOCK

    
    for i in range (len(filtered_links)):
        try:
            code = filtered_links[i]
            #print(code)
            basic_info = scrape_data_with_selenium(driver, code, inpStr, missed)

            if basic_info:
                #print(basic_info)
                if(len(basic_info) < 80):
                    worksheet.insert_row(basic_info[:67], 3)
                else:
                    worksheet.insert_row(basic_info, 3)
                print('-------------FINISHED ' + code + '-------------')
            else:
                print("Failed to scrape data from the given URL.")
        except TimeoutException:
            print(f"Timeout occurred while processing link: {code}. Moving on to the next link.")
            continue
        except Exception as e:
            print(f"Error occurred while processing link CALENDAR: {code}\nError: {e}")
            continue
    

def login_etrader(driver):
    # Open the website
    driver.get("https://us.etrade.com/etx/pxy/login")

    # Find the username, password, and login elements
    username = driver.find_element(By.ID, "USER")
    password = driver.find_element(By.ID, "password")
    login = driver.find_element(By.ID, "mfaLogonButton")

    # Enter your credentials and click the login button
    username.send_keys("garbear04")
    password.send_keys("Garison@@2004")
    login.click()

    time.sleep(20)


def loginCharles(driver):
    #juliusd70

    return

# Main script
userInput = input('1 for Specific Stocks - C for Calendar:  ').strip()
missed = []

with webdriver.Chrome(service=service, options=options) as driver:

    #Login to ETrade 
    #login_etrader(driver)

    #Login to Charles
    #loginCharles(driver) 

    if userInput.isdigit():
        process_specific_stocks(driver, missed)
    else:
        process_calendar(driver, missed)

worksheet2.insert_row(missed, 3)
print("Data inserted into Google Sheets.")