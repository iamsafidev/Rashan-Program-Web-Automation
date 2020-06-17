# Reading an excel file using Python 
import xlrd 
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By

# Give the location of the file 
loc = ("/home/safi/Desktop/tables.xls") 
#Attributes of the Form of Ehsaas Registration
name= ""
cnic=""
mobile=""
house_address=""
house_no=""
city=""
gender_index=""
ms_index=""
#Load the Chrome Driver
driver = webdriver.Chrome()
driver.get("https://rashan.pass.gov.pk/wfbeneficiary.aspx")

# To open Workbook 
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0) 
sheet.cell_value(5,0)


# print(sheet.nrows)
# for i in range(sheet.nrows):
for i in range(sheet.nrows):
    # for j in range(sheet.ncols):
        if((sheet.cell_value(i,1) == 'Name') or (sheet.cell_value(i,0) == 'Registered / UnRegistered')
        or (sheet.cell_value(i,2) == 'Area') or (sheet.cell_value(i,3) == 'City')
        or (sheet.cell_value(i,5) == 'Tehsil') or (sheet.cell_value(i,6) == 'CNIC')
        or (sheet.cell_value(i,7) == 'Mobile') or (sheet.cell_value(i,8) == 'CNIC Issue Date')
        or (sheet.cell_value(i,9) == 'Martial_Status') or (sheet.cell_value(i,10) == 'Gender')
        or (sheet.cell_value(i,11) == 'Flat_No') or (sheet.cell_value(i,4) == 'District')
        ):
            continue
        else:
            name = sheet.cell_value(i,1)
            cnic = sheet.cell_value(i,6)
            mobile=sheet.cell_value(i,7)
            house_address=sheet.cell_value(i,2)
            house_no=sheet.cell_value(i,11)
            city=sheet.cell_value(i,3)
            # ================= GENDER ============
            gender=sheet.cell_value(i,10)
            if(gender == "Male"):
                gender_index=1
            elif(gender == "Female"):
                gender_index=2
            else:
                gender_index=3
            # ++++++++ MARTIAL STATUS ++++++++++++
            Martial_Status= sheet.cell_value(i,9)
            if(Martial_Status == "Single"):
                ms_index=1
            elif(Martial_Status == "Married"):
                ms_index=2
            elif((Martial_Status == "Widow") or (Martial_Status == "Widower") ):
                ms_index=3
            else:
                ms_index=4
            # ====================================
            day,month,year = sheet.cell_value(i,8).split('/')

            # For Loop End Here
            print("I am Outside For Loop")
            print("Day"+day+"Month"+month+"Year"+year)
            fullnamebox = driver.find_element_by_xpath('//*[@id="txtName"]')

            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="txtName"]'))).clear()
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="txtName"]'))).send_keys(name)
            # fullnamebox.send_keys(name)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="txtcnic"]'))).clear()
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="txtcnic"]'))).send_keys(cnic.replace('-', ''))
            
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="txtMobileNo"]'))).clear()
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="txtMobileNo"]'))).send_keys(mobile.replace('-', ''))
            
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="txtHouseNo"]'))).clear()
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="txtHouseNo"]'))).send_keys(house_no)
            
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="txtAddress"]'))).clear()
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="txtAddress"]'))).send_keys(house_address)
            
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="txtcityvillage"]'))).clear()
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="txtcityvillage"]'))).send_keys(city)
            
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//select[@name='txtbcnicissueDay']/option[text()="+month+"]"))).click()
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//select[@name='txtbcnicissueMonth']/option[text()="+day+"]"))).click()
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//select[@name='txtbcnicissueYear']/option[text()="+year+"]"))).click()
            
            select = Select(driver.find_element_by_id('ddGender'))
            select.select_by_index(gender_index)
            driver.implicitly_wait(2)
            select = Select(driver.find_element_by_id('ddMaritalStatus'))
            select.select_by_index(ms_index)
            driver.implicitly_wait(10)
            if(not(driver.find_element_by_xpath('//*[@id="ckUse_My_Particulars"]').is_selected())):
                driver.find_element_by_xpath('//*[@id="ckUse_My_Particulars"]').click()
            driver.find_element_by_xpath("//select[@name='ddDistrict1']/option[text()='HARIPUR']").click()
            driver.implicitly_wait(10)
            driver.find_element_by_xpath("//select[@name='ddTehsil1']/option[text()='HARIPUR']").click()
            #=============++Code FOR +++++++++++++++++++++++++++++
            driver.find_element_by_xpath('//*[@id="bnSave"]').click()
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//button[text()='Close']"))).click()
           