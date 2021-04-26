from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import xlsxwriter
import time

web = webdriver.Chrome('C:\Program Files\Google\Chrome\Application\chromedriver.exe')
#ضع الرابط في الأسفل
web.get('')
workbook = xlsxwriter.Workbook('درجات دفعتي.xlsx')
worksheet = workbook.add_worksheet()

worksheet.write('A1', 'رقم الجلوس')
worksheet.write('B1', 'الإسم')
worksheet.write('C1', 'تصميم المنشآت الخرسانية (5)')
worksheet.write('D1', 'الهندسة الصحية والبيئية (1)')
worksheet.write('E1', 'هندسة الطرق والمطارات')
worksheet.write('F1', 'هندسة السكك الحديدية')
worksheet.write('G1', 'هندسة الموانئ وحماية الشواطئ')
worksheet.write('H1', 'مقرر اختيارى (3) م')
worksheet.write('I1', 'هندسة الملاحة الداخليه')

element = web.find_element_by_xpath('//*[@id="aspnetForm"]')

time.sleep(1)

FacDropDownList = web.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_FacDropDownList"]')
FacDropDownList.click()

time.sleep(1)

option9 = web.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_FacDropDownList"]/option[11]')
option9.click()

time.sleep(1)

PhaseDropDownList = web.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_PhaseDropDownList"]')
PhaseDropDownList.click()

time.sleep(1)

option7 = web.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_PhaseDropDownList"]/option[7]')
option7.click()

time.sleep(1)

DepDropDownList = web.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_DepDropDownList"]')
DepDropDownList.click()

time.sleep(1)

option6 = web.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_DepDropDownList"]/option[6]')
option6.click()

time.sleep(1)


#أول رقم جلوس هنا
Num = 

A = 2
B = 2
C = 2
D = 2
E = 2
F = 2
G = 2
H = 2
I = 2

try:
    #  أخر رقم جلوس + 1
    while Num < :
        SeatNumTextBox = web.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_SeatNumTextBox"]')
        # time.sleep(1)

        SeatNumTextBox.clear()
        # time.sleep(1)

        SeatNumTextBox.send_keys(Num)

        SearchButton = web.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_SearchButton"]')
        SearchButton.click()

        # time.sleep(1)

        Name = web.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_ResPanel"]/table/tbody/tr[2]/td[1]')
        worksheet.write('B' + str(B), Name.text)

        Button1 = web.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_StudentsRep_ctl00_Button1"]')
        Button1.click()

        Concrete = web.find_element_by_xpath('//*[@id="aspnetForm"]/div[3]/div[2]/div/div[1]/table/tbody/tr[2]/td[2]')
        Sanitary = web.find_element_by_xpath('//*[@id="aspnetForm"]/div[3]/div[2]/div/div[1]/table/tbody/tr[3]/td[2]')
        Road = web.find_element_by_xpath('//*[@id="aspnetForm"]/div[3]/div[2]/div/div[1]/table/tbody/tr[4]/td[2]')
        Railway = web.find_element_by_xpath('//*[@id="aspnetForm"]/div[3]/div[2]/div/div[1]/table/tbody/tr[5]/td[2]')
        Ports = web.find_element_by_xpath('//*[@id="aspnetForm"]/div[3]/div[2]/div/div[1]/table/tbody/tr[6]/td[2]')
        Course = web.find_element_by_xpath('//*[@id="aspnetForm"]/div[3]/div[2]/div/div[1]/table/tbody/tr[7]/td[2]')
        Navigation = web.find_element_by_xpath('//*[@id="aspnetForm"]/div[3]/div[2]/div/div[1]/table/tbody/tr[8]/td[2]')

        worksheet.write('A' + str(A), Num)
        worksheet.write('C' + str(C), Concrete.text)
        worksheet.write('D' + str(D), Sanitary.text)
        worksheet.write('E' + str(E), Road.text)
        worksheet.write('F' + str(F), Railway.text)
        worksheet.write('G' + str(G), Ports.text)
        worksheet.write('H' + str(H), Course.text)
        worksheet.write('I' + str(I), Navigation.text)

        web.back()

        Num = Num + 1

        A += 1
        B += 1
        C += 1
        D += 1
        E += 1
        F += 1
        G += 1
        H += 1
        I += 1

except:
    web.back()

    Num = Num + 1

    A += 1
    B += 1
    C += 1
    D += 1
    E += 1
    F += 1
    G += 1
    H += 1
    I += 1

    while Num < 47125:
        SeatNumTextBox = web.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_SeatNumTextBox"]')
        # time.sleep(1)

        SeatNumTextBox.clear()
        # time.sleep(1)

        SeatNumTextBox.send_keys(Num)

        SearchButton = web.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_SearchButton"]')
        SearchButton.click()

        # time.sleep(1)

        Name = web.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_ResPanel"]/table/tbody/tr[2]/td[1]')
        worksheet.write('B' + str(B), Name.text)

        Button1 = web.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_StudentsRep_ctl00_Button1"]')
        Button1.click()

       

        Concrete = web.find_element_by_xpath('//*[@id="aspnetForm"]/div[3]/div[2]/div/div[1]/table/tbody/tr[2]/td[2]')
        Sanitary = web.find_element_by_xpath('//*[@id="aspnetForm"]/div[3]/div[2]/div/div[1]/table/tbody/tr[3]/td[2]')
        Road = web.find_element_by_xpath('//*[@id="aspnetForm"]/div[3]/div[2]/div/div[1]/table/tbody/tr[4]/td[2]')
        Railway = web.find_element_by_xpath('//*[@id="aspnetForm"]/div[3]/div[2]/div/div[1]/table/tbody/tr[5]/td[2]')
        Ports = web.find_element_by_xpath('//*[@id="aspnetForm"]/div[3]/div[2]/div/div[1]/table/tbody/tr[6]/td[2]')
        Course = web.find_element_by_xpath('//*[@id="aspnetForm"]/div[3]/div[2]/div/div[1]/table/tbody/tr[7]/td[2]')
        Navigation = web.find_element_by_xpath('//*[@id="aspnetForm"]/div[3]/div[2]/div/div[1]/table/tbody/tr[8]/td[2]')

        worksheet.write('A' + str(A), Num)
        worksheet.write('C' + str(C), Concrete.text)
        worksheet.write('D' + str(D), Sanitary.text)
        worksheet.write('E' + str(E), Road.text)
        worksheet.write('F' + str(F), Railway.text)
        worksheet.write('G' + str(G), Ports.text)
        worksheet.write('H' + str(H), Course.text)
        worksheet.write('I' + str(I), Navigation.text)

        web.back()

        Num = Num + 1

        A += 1
        B += 1
        C += 1
        D += 1
        E += 1
        F += 1
        G += 1
        H += 1
        I += 1

workbook.close()
