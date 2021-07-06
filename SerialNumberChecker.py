# Programmer: rip97
# File: SerialNumberChecker.py
# Date: 9/23/19
# Description: Program takes use of Selenium API and Beautiful Soup Library to automate process
# of looking up serial numbers for HP computers to check warranty status.


import time
from selenium import webdriver
from selenium.common import exceptions
from bs4 import BeautifulSoup
import csv


def GetWarranty(serialNumber):
    driver = webdriver.Chrome(executable_path="chromedriver.exe")
    try:

        driver.implicitly_wait(5)
        driver.get("https://support.hp.com/us-en/checkwarranty")

        textBox = driver.find_element_by_id("wFormSerialNumber")
        textBox.send_keys(serialNumber)
        btn = driver.find_element_by_id("btnWFormSubmit")
        btn.click()

        # switch to next url after from submit
        for handle in driver.window_handles:
            driver.switch_to.window(handle)

        time.sleep(2)  # thank you stackOverFlow for this line code to make this work!!!

        return driver

    except exceptions as e:
        print("Error occurred: %s" % e)


def ScrapePage(driver):
    soup = BeautifulSoup(driver.page_source, "html.parser")
    driver.quit()

    result = soup.find_all(class_="col-lg-17")
    warrantyResult = []

    counter = 0
    for line in result:
        line = str(line)
        line = line.split()
        line.pop(1)
        line.pop(0)
        line.pop(-1)

        if counter == 0:
            s = " "
            s = s.join(line)
            warrantyResult.append(s)
        if counter == 1:
            s = " "
            s = s.join(line)
            warrantyResult.append(s)
        if counter == 2:
            s = " "
            line.pop(0)
            if len(line) == 1:
                s = s.join(line)
                s = s.strip('style="color:#FF0000;">')
                s = s.strip('</span>')
                warrantyResult.append(s)
            else:
                s = s.join(line[1:8])
                warrantyResult.append(s)
        if counter == 3:
            s = " "
            s = s.join(line)
            warrantyResult.append(s)

        if counter == 4:
            s = " "
            s = s.join(line)
            warrantyResult.append(s)
        counter += 1
    return warrantyResult


def FormatData(results):
    title = "*** Warranty Information ***\n" \
            " Warranty Type: %s\n" \
            " Service Type: %s\n" \
            " Status: %s\n" \
            " Start Date: %s\n" \
            " End Date: %s " % (results[0], results[1], results[2], results[3], results[4])
    print(title)


def WriteData(results, fileContent):
    headers = ['AssetTag', 'SerialNumber', 'End Date', 'Start Date', 'Status', 'Warranty Type', 'Warranty']

    counter = 0
    for i in results:
        i.append(fileContent[counter][1])
        i.append(fileContent[counter][0])
        counter += 1

    with open('Warranty_Results.csv', 'w', newline='') as csvF:
        writer = csv.writer(csvF)
        writer.writerow(headers)
        counter = 0
        for result in results:
            writer.writerow(result[::-1])
            counter += 1
    csvF.close()


def main():

    doAnother = 'y'
    while doAnother == 'y':
        print("Do you want to check one serial number or many serial numbers?\n"
              "Press (1) for a single serial number.\n"
              "Press (2) for multiple serial numbers\n")
        option = int(input("-->"))
        if option == 1:
            num = str(input("Please Enter a Serial Number -->")).upper()
            driver = GetWarranty(num)
            results = ScrapePage(driver)
            time.sleep(1)
            FormatData(results)

            doAnother = str(input("\nWould you like to check another?\n"
                                  "Press 'y' for yes.\n"
                                  "To exit, press any key then hit enter.\n"
                                  "-->")).lower()

        else:
            filePath = str(input("Please make sure the serial numbers are in a text file.\n"
                                 "Enter the text files path along with the file: "))
            try:
                serialNumbers = []
                print(f'Reading File . . . ')
                file = open(filePath, 'r')
                for line in file:
                    line = str(line).strip()
                    line = str(line).split()
                    if line:
                        serialNumbers.append(line)
                file.close()

                print(f'Checking Warranties . . . ')
                warResults = []
                counter = 0

                for numbers in serialNumbers:
                    if counter == len(serialNumbers) // 2:
                        print("Program sleeping for 60 secs")
                        time.sleep(60)
                    driver = GetWarranty(numbers[1])
                    results = ScrapePage(driver)
                    warResults.append(results)
                    print(f"Completed Search of {counter + 1} out of {len(serialNumbers)}")
                    counter += 1
                WriteData(warResults, serialNumbers)
            except exceptions as e:
                print(f'File Error: {e}')

            print("Check Directory for output file")
            doAnother = 'n'


if __name__ == "__main__":
    main()
