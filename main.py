import xlsxwriter
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions
from webdriver_manager.chrome import ChromeDriverManager


def CreateExcel(dataMatrix):
    workbook = xlsxwriter.Workbook("data.xlsx")
    worksheet = workbook.add_worksheet()
    worksheet.write(0, 0, "Name")
    worksheet.write(0, 1, "Rating")
    worksheet.write(0, 2, "Reviews")
    worksheet.write(0, 3, "Phone no")
    worksheet.write(0, 4, "Link")
    worksheet.write(0, 5, "Address")
    for i in range(len(dataMatrix)):
        for j in range(len(dataMatrix[0])):
            worksheet.write(i + 1, j, str(dataMatrix[i][j]))

    workbook.close()


def get_details(comp, wait):
    wait.until(expected_conditions.visibility_of_element_located((By.CLASS_NAME, "bfIbhd")))
    details = driver.find_elements(By.CLASS_NAME, "bfIbhd")
    name = comp.find_elements(By.CLASS_NAME, "rgnuSb")[0].text
    try:
        rating = comp.find_elements(By.CLASS_NAME, "rGaJuf")[0].text
    except IndexError:
        rating = "None"
    try:
        reviews = comp.find_elements(By.CLASS_NAME, "leIgTe")[0].text
    except IndexError:
        reviews = "None"
    try:
        phone_no = details[0].find_elements(By.CLASS_NAME, "eigqqc")[0].text
    except IndexError:
        phone_no = "None"
    try:
        link = details[0].find_elements(By.CLASS_NAME, "Gx8NHe")[0].text
    except IndexError:
        link = "None"
    try:
        addr = details[0].find_elements(By.CLASS_NAME, "hgRN0")[0].text
    except IndexError:
        addr = "None"

    return [name, rating, reviews, phone_no, link, addr]


def parseInfo(driver, wait):
    dataMatrix = []
    companies = driver.find_elements(By.CLASS_NAME, "DVBRsc")

    for comp in companies:
        comp.click()
        data = get_details(comp, wait)
        dataMatrix.append(data)


    return dataMatrix


if __name__ == "__main__":
    # -------------- setup driver ------------
    url = "https://www.google.com/localservices/prolist?g2lbs=ANTchaMTaYdPs4HtAbU7RKqm60615Q-wNOs6Xd-4vKCXyF7ILae6aBjWOK0yKTAff0jCt9oRPlvy25r91MqOczGrndcqyy4uLONLJhWw0CG_mtnoQ_OgTAgFqAhLPmgC6pksffG_sXw1&hl=en-IN&gl=in&ssta=1&q=mobile%20detailing%20atlanta%20ga&oq=mobile%20detailing%20atlanta%20ga&slp=MgA6HENoTUk5b0hMMEpIZWdRTVYtaEo3QngxYldnRDhSAggCYACSAbcCCg0vZy8xMWd4bnF4aHhsCg0vZy8xMWZ3a25yZjF2Cg0vZy8xMWI2ZHZrNnF0Cg0vZy8xMWdqX204MmhmCg0vZy8xMWp3anQwczJuCg0vZy8xMWd3Mnh0NTNfCg0vZy8xMWpoN3FoNW5xCg0vZy8xMWYzemZiMGowCg0vZy8xMXJyYjRxOGNqCg0vZy8xMWNuOXNiNHhsCgwvZy8xcHAydmNmcmIKDS9nLzExZzZfMTAwcjIKDS9nLzExdGNoa18ycTMKDS9nLzExc2M5cjd0cF8KDS9nLzExc2p0M2doN2cKDS9nLzExZm01NHQxOG4KDS9nLzExbHB6bnFoYjMKDS9nLzExY20wNjFoc2wKDS9nLzExZ2hzZzl0dDkKDS9nLzExcHpyczN2azASBBICCAESBAoCCAGaAQYKAhcZEAA%3D&src=2&serdesk=1&sa=X&ved=2ahUKEwi6osXQkd6BAxU0YvUHHZFLCBAQjGp6BAgTEAE&scp=ChpnY2lkOmNhcl9kZXRhaWxpbmdfc2VydmljZRJSEhIJjQmTaV0E9YgRC2MLmS_e_mYaEgnR02eYcKv1iBETxdcOi0HS1yIQQXRsYW50YSwgR0EsIFVTQSoUDThEDhQVhZSazR1etzIUJchwws0wABoQbW9iaWxlIGRldGFpbGluZyIbbW9iaWxlIGRldGFpbGluZyBhdGxhbnRhIGdhKhBWYWxldGluZyBzZXJ2aWNlOj0KCS9tLzAzemowZxIaam9iX3R5cGVfaWQ6YXV0b19kZXRhaWxpbmcaDkF1dG8gZGV0YWlsaW5nIgJlbjAB"
    options = Options()
    options.add_experimental_option("detach", True)
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.get(url)
    wait = WebDriverWait(driver, timeout=15)
    # -----------------------------------------

    dataMatrix = []
    endreached = False
    while not endreached:
        try:
            wait.until(expected_conditions.presence_of_element_located((By.XPATH, '//*[@id="yDmH0d"]//*[text()[contains(.,"Next >")]]')))
            button = driver.find_element(By.XPATH, '//*[@id="yDmH0d"]//*[text()[contains(.,"Next >")]]') 
            dataMatrix.extend(parseInfo(driver, wait))
            button.click()
        except:
            dataMatrix.extend(parseInfo(driver, wait))
            print("-end-")
            endreached = True

    print("creating Excel...")
    CreateExcel(dataMatrix)
    print("[ OK ] Excel file Created!")
