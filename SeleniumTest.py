import time
from selenium import webdriver
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from openpyxl import load_workbook

def save_data_excel(content):  
	workbook_name = 'Sample_data.xlsx'
	wb = load_workbook(workbook_name)
	page = wb.active
	page.append(content)
	wb.save(filename=workbook_name)

# write your pids here to search with
pids = ["002-630-141", "002-630-141"]

driver = webdriver.Firefox()

driver.set_page_load_timeout(1000) #1000 seconds
url = "https://www.bcassessment.ca/Property/AssessmentSearch?bcalogin=1"
driver.get(url)

for pid in pids:
    select_fr = Select(driver.find_element_by_id("ddlSearchType"))
    select_fr.select_by_index(3)

    data = []

    driver.find_element_by_name("searchPID").send_keys(pid)
    driver.find_element_by_id("btnSearch").click()

    #time.sleep(20)
    
    delay = 100 # seconds
    try:
        myElem = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.ID, 'mainaddresstitle')))
        print("Page is ready!")
    except TimeoutException:
        print("Loading took too much time!")
    

    address = driver.find_element_by_css_selector('#mainaddresstitle').get_attribute('textContent')
    areajursrolltext = driver.find_element_by_id("areajursrolltitlebox").get_attribute('textContent')
    pos = areajursrolltext.find(":")
    areajursroll = areajursrolltext[pos+1:len(areajursrolltext)].strip()
    totalVal = driver.find_element_by_class_name("valuesection").get_attribute('textContent')
    assesment = driver.find_element_by_xpath('//*[@id="lblLastAssessmentDate"]').get_attribute('textContent')
    land = driver.find_element_by_xpath('//*[@id="lblTotalAssessedLand"]').get_attribute('textContent')
    building = driver.find_element_by_xpath('//*[@id="lblTotalAssessedBuilding"]').get_attribute('textContent')
    prev_year_value = driver.find_element_by_xpath('//*[@id="lblPreviousAssessedValue"]').get_attribute('textContent')
    prev_year_land = driver.find_element_by_xpath('//*[@id="lblPreviousAssessedLand"]').get_attribute('textContent')
    prev_year_building = driver.find_element_by_xpath('//*[@id="lblPreviousAssessedBuilding"]').get_attribute('textContent')
    year_built = driver.find_element_by_xpath('//*[@id="lblYearBuilt"]').get_attribute('textContent')
    description = driver.find_element_by_xpath('//*[@id="lblDescription"]').get_attribute('textContent')

    bedroom = driver.find_element_by_xpath('//*[@id="lblBedrooms"]').get_attribute('textContent')
    bath = driver.find_element_by_xpath('//*[@id="lblBathRooms"]').get_attribute('textContent')
    carpet = driver.find_element_by_xpath('//*[@id="lblCarPorts"]').get_attribute('textContent')
    garage = driver.find_element_by_xpath('//*[@id="lblGarages"]').get_attribute('textContent')
    land_size = driver.find_element_by_xpath('//*[@id="lblLandSize"]').get_attribute('textContent')
    first_floor_area = driver.find_element_by_xpath('//*[@id="lblFirstFloorArea"]').get_attribute('textContent')
    second_floor_area = driver.find_element_by_xpath('//*[@id="lblSecondFloorArea"]').get_attribute('textContent')
    basement_finish_area = driver.find_element_by_xpath('//*[@id="lblBasementFinishArea"]').get_attribute('textContent')
    strata_area = driver.find_element_by_xpath('//*[@id="lblStrataTotalArea"]').get_attribute('textContent')
    building_storeys = driver.find_element_by_xpath('//*[@id="lblStoriesBuilding"]').get_attribute('textContent')
    gross_leasable_area = driver.find_element_by_xpath('//*[@id="lblGrossLeasableArea"]').get_attribute('textContent')
    net_leasable_area = driver.find_element_by_xpath('//*[@id="lblNetLeasableArea"]').get_attribute('textContent')
    no_apartment_units = driver.find_element_by_xpath('//*[@id="lblNumberUnitApartment"]').get_attribute('textContent')

    legal_description = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div[2]/div[4]/div/div[2]/div[1]/div/p[1]').get_attribute('textContent')
    parcel_id = pid
    manufa_home = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div[2]/div[4]/div/div[2]/p[3]').get_attribute('textContent')
    width = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div[2]/div[4]/div/div[2]/div[3]/div[1]/div[1]').get_attribute('textContent')
    length = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div[2]/div[4]/div/div[2]/div[3]/div[2]/div[1]').get_attribute('textContent')
    total_area = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div[2]/div[4]/div/div[2]/div[3]/div[3]/div[1]').get_attribute('textContent')


    #printing all values
    print(pid)
    print(address)
    print(areajursroll)
    print(totalVal)
    print(assesment)
    print(land)
    print(building)
    print(prev_year_value)
    print(prev_year_land)
    print(prev_year_building)
    print(year_built)
    print(description)
    print(bedroom)
    print(bath)
    print(carpet)
    print(garage)
    print(land_size)
    print(first_floor_area)
    print(second_floor_area)
    print(basement_finish_area)
    print(strata_area)
    print(building_storeys)
    print(gross_leasable_area)
    print(net_leasable_area)
    print(no_apartment_units)
    print(legal_description)
    print(parcel_id)
    print(manufa_home)
    print(width)
    print(length)
    print(total_area)

    # append this data to data list
    data.append(pid)
    data.append(address)
    data.append(areajursroll)
    data.append(totalVal)
    data.append(assesment)
    data.append(land)
    data.append(building)
    data.append(prev_year_value)
    data.append(prev_year_land)
    data.append(prev_year_building)
    data.append(year_built)
    data.append(description)
    data.append(bedroom)
    data.append(bath)
    data.append(carpet)
    data.append(garage)
    data.append(land_size)
    data.append(first_floor_area)
    data.append(second_floor_area)
    data.append(basement_finish_area)
    data.append(strata_area)
    data.append(building_storeys)
    data.append(gross_leasable_area)
    data.append(net_leasable_area)
    data.append(no_apartment_units)
    data.append(legal_description)
    data.append(parcel_id)
    data.append(manufa_home)
    data.append(width)
    data.append(length)
    data.append(total_area)

    # save data to excel
    save_data_excel(data)

time.sleep(10)
driver.quit()