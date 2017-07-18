from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import time
import csv
import sys
import re
import os
import wmi
import random


def parsePage(catalogName, driver, city, path):

    time.sleep(5)
    print(city)
    driver.find_element_by_xpath("//a[@name='" + catalogName + " " + city + "']").click()
    time.sleep(5)
    driver.find_element_by_xpath(".//search-result-list/ol//li[1]//div/div[1]/div[1]/h2/a/span").click()
    time.sleep(5)

    time.sleep(0.6)
    inputElement = driver.find_element_by_xpath(".//*[@id='entry-detail']/section/div/div[1]/div[2]/span/span[3]")
    numberOfPage = int(inputElement.text)
    print(numberOfPage)
    f = open(path[0] + '/' + catalogName + '.csv', 'a', encoding="utf-8")

    for i in range(numberOfPage):
        if i % 20 == 0:
            # Obtain network adaptors configurations
            nic_configs = wmi.WMI().Win32_NetworkAdapterConfiguration(IPEnabled=True)
            # First network adaptor
            nic = nic_configs[0]
            # IP address, subnetmask and gateway values should be unicode objects
            ip = '192.168.0.' + str(random.randint(1, 99))
            subnetmask = '255.255.255.0'
            gateway = '192.168.0.' + str(random.randint(1, 99))
            # Set IP address, subnetmask and default gateway
            # Note: EnableStatic() and SetGateways() methods require *lists* of values to be passed
            nic.EnableStatic(IPAddress=[ip], SubnetMask=[subnetmask])
            nic.SetGateways(DefaultIPGateway=[gateway])
        print(i)
        pageNum = str(i+1)
        time.sleep(0.001)

        f.write('"' + catalogName + '"' + ';')

        try:
            inputElement = driver.find_element_by_xpath(".//*[@id='entry']/div/div[2]/div[1]/h1")
            name = inputElement.text
        except:
            name = ""
        f.write('"' + name + '"' + ';')

        time.sleep(0.001)
        try:
            inputElement = driver.find_element_by_xpath(".//*[@id='entry']/div/div[2]/div[3]/div[1]/p/span[1]/span")
            dataSpan1 = inputElement.text
            listOfNumbers = re.findall('\d+', dataSpan1)
            lenghtNumbers = len(listOfNumbers)
            lastNumber = listOfNumbers[lenghtNumbers - 1]
            startNumberOfHome = dataSpan1.find(lastNumber)
            street = dataSpan1[0:startNumberOfHome].lstrip(' ').rstrip(' ')
            numberOfHome = dataSpan1[startNumberOfHome:len(dataSpan1)].lstrip(' ').rstrip(' ')
            if lenghtNumbers > 1:
                preLastNumber = listOfNumbers[lenghtNumbers - 2]
                startPositionPreLastNumberOfHome = dataSpan1.find(preLastNumber)
                LastPositionOfPreLastNumberOfHome = startPositionPreLastNumberOfHome + len(str(preLastNumber))
                startNumberOfHome = dataSpan1.find(lastNumber)
                valueBetweenLastAndPreLastNumber = dataSpan1[LastPositionOfPreLastNumberOfHome:startNumberOfHome]
                if valueBetweenLastAndPreLastNumber.lstrip(' ').rstrip(' ') == '-':
                    street = dataSpan1[0:startPositionPreLastNumberOfHome].lstrip(' ').rstrip(' ')
                    numberOfHome = dataSpan1[startPositionPreLastNumberOfHome:len(dataSpan1)].lstrip(' ').rstrip(' ')
        except:
            street =""
            numberOfHome = ""
        f.write('"' + street + '"' + ';')
        f.write('"' + numberOfHome + '"' + ';')

        time.sleep(0.001)
        try:
            inputElement = driver.find_element_by_xpath(".//*[@id='entry']/div/div[2]/div[3]/div[1]/p/span[2]/span[1]")
            postcode = inputElement.text
        except:
            postcode = ""
        f.write('"' + postcode + '"' + ';')

        time.sleep(0.001)
        try:
            inputElement = driver.find_element_by_xpath(".//*[@id='entry']/div/div[2]/div[3]/div[1]/p/span[2]")
            dataSpan2 = inputElement.text
            city = dataSpan2[len(postcode):].lstrip(' ').rstrip(' ')
        except:
            city = ""
        f.write('"' + city + '"' + ';')
        f.write('"Deutschland"' + ';')

        time.sleep(0.001)
        lastRightUrl = driver.current_url
        try:
            driver.find_element_by_xpath("//span[text()='Telefonnummer, Fax, Mobil']/..//span[text()='Kontakt anzeigen!']/..").click()
        except:
            print()

        try:
            time.sleep(1)
            inputElement = driver.find_element_by_xpath("//span[text()='Telefonnummer, Fax, Mobil']/../div[1]/div[1]/div[1]/div[1]/p")
            dataTelephone1 = inputElement.text
            if dataTelephone1 == 'Telefonnummer':
                dataTelephone1 = ""
            if dataTelephone1 == 'Telefonn...':
                dataTelephone1 = ""
        except:
            dataTelephone1 = ""
        f.write('"' + dataTelephone1 + '"' + ';')

        try:
            time.sleep(0.001)
            inputElement = driver.find_element_by_xpath("//span[text()='Telefonnummer, Fax, Mobil']/../div[1]/div[1]/div[1]/div[2]/p")
            dataTelephone2 = inputElement.text
            if dataTelephone2 == 'Faxnummer':
                dataTelephone2 = ""
            if dataTelephone2 == 'Faxnum...':
                dataTelephone2 = ""
        except:
            dataTelephone2 = ""
        f.write('"' + dataTelephone2 + '"' + ';')

        try:
            time.sleep(0.001)
            inputElement = driver.find_element_by_xpath("//span[text()='Telefonnummer, Fax, Mobil']/../div[1]/div[1]/div[1]/div[3]/p")
            dataTelephone3 = inputElement.text
            if dataTelephone3 == 'Mobilnummer':
                dataTelephone3 = ""
            if dataTelephone3 == 'Mobilnu...':
                dataTelephone3 = ""
        except:
            dataTelephone3 = ""
        f.write('"' + dataTelephone3 + '"' + ';')

        try:
            time.sleep(0.001)
            inputElement = driver.find_element_by_xpath("//span[text()='Telefonnummer, Fax, Mobil']/..//span[text()='zur Website']/..")
            href = inputElement.get_attribute("href")
        except:
            href = ""
        f.write('"' + href + '"')

        time.sleep(0.001)

        f.write('\n')

        time.sleep(0.001)

        if i == 0:
            driver.find_element_by_xpath(".//*[@id='entry-detail']/section/div/div[1]/div[2]/span/a").click()
        elif i < numberOfPage:
            try:
                inputElement = driver.find_element_by_xpath(
                    ".//*[@id='entry-detail']/section/div/div[1]/div[2]/span/a[2]").click()
            except:
                print('End cycle')
        else:
            break
        time.sleep(0.2)

    time.sleep(0.1)
    f.close()


def parsePageWithCatalogId(catalogName, driver, city, number, path):

    time.sleep(5)
    print(city)
    driver.find_element_by_xpath(".//div/div/form/div/div[2]/div/a[" + number + "]").click()

    time.sleep(5)
    driver.find_element_by_xpath(".//search-result-list/ol//li[1]//div/div[1]/div[1]/h2/a/span").click()
    time.sleep(5)

    time.sleep(0.6)
    inputElement = driver.find_element_by_xpath(".//*[@id='entry-detail']/section/div/div[1]/div[2]/span/span[3]")
    numberOfPage = int(inputElement.text)
    print(numberOfPage)
    f = open(path[0] + '/' + catalogName + '.csv', 'a', encoding="utf-8")

    for i in range(numberOfPage):
        if i % 20 == 0:
            # Obtain network adaptors configurations
            nic_configs = wmi.WMI().Win32_NetworkAdapterConfiguration(IPEnabled=True)
            # First network adaptor
            nic = nic_configs[0]
            # IP address, subnetmask and gateway values should be unicode objects
            ip = '192.168.0.' + str(random.randint(1, 99))
            subnetmask = '255.255.255.0'
            gateway = '192.168.0.' + str(random.randint(1, 99))
            # Set IP address, subnetmask and default gateway
            # Note: EnableStatic() and SetGateways() methods require *lists* of values to be passed
            nic.EnableStatic(IPAddress=[ip], SubnetMask=[subnetmask])
            nic.SetGateways(DefaultIPGateway=[gateway])
        print(i)
        pageNum = str(i+1)
        time.sleep(0.001)

        f.write('"' + catalogName + '"' + ';')

        try:
            inputElement = driver.find_element_by_xpath(".//*[@id='entry']/div/div[2]/div[1]/h1")
            name = inputElement.text
        except:
            name = ""
        f.write('"' + name + '"' + ';')

        time.sleep(0.001)
        try:
            inputElement = driver.find_element_by_xpath(".//*[@id='entry']/div/div[2]/div[3]/div[1]/p/span[1]/span")
            dataSpan1 = inputElement.text
            listOfNumbers = re.findall('\d+', dataSpan1)
            lenghtNumbers = len(listOfNumbers)
            lastNumber = listOfNumbers[lenghtNumbers - 1]
            startNumberOfHome = dataSpan1.find(lastNumber)
            street = dataSpan1[0:startNumberOfHome].lstrip(' ').rstrip(' ')
            numberOfHome = dataSpan1[startNumberOfHome:len(dataSpan1)].lstrip(' ').rstrip(' ')
            if lenghtNumbers > 1:
                preLastNumber = listOfNumbers[lenghtNumbers - 2]
                startPositionPreLastNumberOfHome = dataSpan1.find(preLastNumber)
                LastPositionOfPreLastNumberOfHome = startPositionPreLastNumberOfHome + len(str(preLastNumber))
                startNumberOfHome = dataSpan1.find(lastNumber)
                valueBetweenLastAndPreLastNumber = dataSpan1[LastPositionOfPreLastNumberOfHome:startNumberOfHome]
                if valueBetweenLastAndPreLastNumber.lstrip(' ').rstrip(' ') == '-':
                    street = dataSpan1[0:startPositionPreLastNumberOfHome].lstrip(' ').rstrip(' ')
                    numberOfHome = dataSpan1[startPositionPreLastNumberOfHome:len(dataSpan1)].lstrip(' ').rstrip(' ')
        except:
            street =""
            numberOfHome = ""
        f.write('"' + street + '"' + ';')
        f.write('"' + numberOfHome + '"' + ';')

        time.sleep(0.001)
        try:
            inputElement = driver.find_element_by_xpath(".//*[@id='entry']/div/div[2]/div[3]/div[1]/p/span[2]/span[1]")
            postcode = inputElement.text
        except:
            postcode = ""
        f.write('"' + postcode + '"' + ';')

        time.sleep(0.001)
        try:
            inputElement = driver.find_element_by_xpath(".//*[@id='entry']/div/div[2]/div[3]/div[1]/p/span[2]")
            dataSpan2 = inputElement.text
            city = dataSpan2[len(postcode):].lstrip(' ').rstrip(' ')
        except:
            city = ""
        f.write('"' + city + '"' + ';')
        f.write('"Deutschland"' + ';')

        time.sleep(0.001)
        lastRightUrl = driver.current_url
        try:
            driver.find_element_by_xpath("//span[text()='Telefonnummer, Fax, Mobil']/..//span[text()='Kontakt anzeigen!']/..").click()
        except:
            print()

        try:
            time.sleep(1)
            inputElement = driver.find_element_by_xpath("//span[text()='Telefonnummer, Fax, Mobil']/../div[1]/div[1]/div[1]/div[1]/p")
            dataTelephone1 = inputElement.text
            if dataTelephone1 == 'Telefonnummer':
                dataTelephone1 = ""
            if dataTelephone1 == 'Telefonn...':
                dataTelephone1 = ""
        except:
            dataTelephone1 = ""
        f.write('"' + dataTelephone1 + '"' + ';')

        try:
            time.sleep(0.001)
            inputElement = driver.find_element_by_xpath("//span[text()='Telefonnummer, Fax, Mobil']/../div[1]/div[1]/div[1]/div[2]/p")
            dataTelephone2 = inputElement.text
            if dataTelephone2 == 'Faxnummer':
                dataTelephone2 = ""
            if dataTelephone2 == 'Faxnum...':
                dataTelephone2 = ""
        except:
            dataTelephone2 = ""
        f.write('"' + dataTelephone2 + '"' + ';')

        try:
            time.sleep(0.001)
            inputElement = driver.find_element_by_xpath("//span[text()='Telefonnummer, Fax, Mobil']/../div[1]/div[1]/div[1]/div[3]/p")
            dataTelephone3 = inputElement.text
            if dataTelephone3 == 'Mobilnummer':
                dataTelephone3 = ""
            if dataTelephone3 == 'Mobilnu...':
                dataTelephone3 = ""
        except:
            dataTelephone3 = ""
        f.write('"' + dataTelephone3 + '"' + ';')

        try:
            time.sleep(0.001)
            inputElement = driver.find_element_by_xpath(
                "//span[text()='Telefonnummer, Fax, Mobil']/..//span[text()='zur Website']/..")
            href = inputElement.get_attribute("href")
        except:
            href = ""
        f.write('"' + href + '"')
        time.sleep(0.001)

        f.write('\n')

        time.sleep(0.001)

        if i == 0:
            driver.find_element_by_xpath(".//*[@id='entry-detail']/section/div/div[1]/div[2]/span/a").click()
        elif i < numberOfPage:
            try:
                inputElement = driver.find_element_by_xpath(".//*[@id='entry-detail']/section/div/div[1]/div[2]/span/a[2]").click()
            except:
                print('End cycle')
        else:
            break
        time.sleep(0.2)

    time.sleep(0.1)
    f.close()


def parsePageSlow(catalogName, driver, city, number, path):

    time.sleep(5)
    print(city)
    driver.find_element_by_xpath(".//div/div/form/div/div[2]/div/a[" + number + "]").click()

    time.sleep(5)
    driver.find_element_by_xpath(".//search-result-list/ol//li[1]//div/div[1]/div[1]/h2/a/span").click()
    time.sleep(5)

    time.sleep(2)
    inputElement = driver.find_element_by_xpath(".//*[@id='entry-detail']/section/div/div[1]/div[2]/span/span[3]")
    numberOfPage = int(inputElement.text)
    print(numberOfPage)
    f = open(path[0] + '/' + catalogName + '.csv', 'a', encoding="utf-8")

    for i in range(numberOfPage):
        if i % 20 == 0:
            # Obtain network adaptors configurations
            nic_configs = wmi.WMI().Win32_NetworkAdapterConfiguration(IPEnabled=True)
            # First network adaptor
            nic = nic_configs[0]
            # IP address, subnetmask and gateway values should be unicode objects
            ip = '192.168.0.' + str(random.randint(1, 99))
            subnetmask = '255.255.255.0'
            gateway = '192.168.0.' + str(random.randint(1, 99))
            # Set IP address, subnetmask and default gateway
            # Note: EnableStatic() and SetGateways() methods require *lists* of values to be passed
            nic.EnableStatic(IPAddress=[ip], SubnetMask=[subnetmask])
            nic.SetGateways(DefaultIPGateway=[gateway])
        print(i)
        pageNum = str(i+1)
        time.sleep(0.001)

        f.write('"' + catalogName + '"' + ';')

        try:
            inputElement = driver.find_element_by_xpath(".//*[@id='entry']/div/div[2]/div[1]/h1")
            name = inputElement.text
        except:
            name = ""
        f.write('"' + name + '"' + ';')

        time.sleep(0.001)
        try:
            inputElement = driver.find_element_by_xpath(".//*[@id='entry']/div/div[2]/div[3]/div[1]/p/span[1]/span")
            dataSpan1 = inputElement.text
            listOfNumbers = re.findall('\d+', dataSpan1)
            lenghtNumbers = len(listOfNumbers)
            lastNumber = listOfNumbers[lenghtNumbers - 1]
            startNumberOfHome = dataSpan1.find(lastNumber)
            street = dataSpan1[0:startNumberOfHome].lstrip(' ').rstrip(' ')
            numberOfHome = dataSpan1[startNumberOfHome:len(dataSpan1)].lstrip(' ').rstrip(' ')
            if lenghtNumbers > 1:
                preLastNumber = listOfNumbers[lenghtNumbers - 2]
                startPositionPreLastNumberOfHome = dataSpan1.find(preLastNumber)
                LastPositionOfPreLastNumberOfHome = startPositionPreLastNumberOfHome + len(str(preLastNumber))
                startNumberOfHome = dataSpan1.find(lastNumber)
                valueBetweenLastAndPreLastNumber = dataSpan1[LastPositionOfPreLastNumberOfHome:startNumberOfHome]
                if valueBetweenLastAndPreLastNumber.lstrip(' ').rstrip(' ') == '-':
                    street = dataSpan1[0:startPositionPreLastNumberOfHome].lstrip(' ').rstrip(' ')
                    numberOfHome = dataSpan1[startPositionPreLastNumberOfHome:len(dataSpan1)].lstrip(' ').rstrip(' ')
        except:
            street =""
            numberOfHome = ""
        f.write('"' + street + '"' + ';')
        f.write('"' + numberOfHome + '"' + ';')

        time.sleep(0.001)
        try:
            inputElement = driver.find_element_by_xpath(".//*[@id='entry']/div/div[2]/div[3]/div[1]/p/span[2]/span[1]")
            postcode = inputElement.text
        except:
            postcode = ""
        f.write('"' + postcode + '"' + ';')

        time.sleep(0.001)
        try:
            inputElement = driver.find_element_by_xpath(".//*[@id='entry']/div/div[2]/div[3]/div[1]/p/span[2]")
            dataSpan2 = inputElement.text
            city = dataSpan2[len(postcode):].lstrip(' ').rstrip(' ')
        except:
            city = ""
        f.write('"' + city + '"' + ';')
        f.write('"Deutschland"' + ';')

        time.sleep(0.001)
        lastRightUrl = driver.current_url
        try:
            driver.find_element_by_xpath("//span[text()='Telefonnummer, Fax, Mobil']/..//span[text()='Kontakt anzeigen!']/..").click()
        except:
            print()

        try:
            time.sleep(2)
            inputElement = driver.find_element_by_xpath("//span[text()='Telefonnummer, Fax, Mobil']/../div[1]/div[1]/div[1]/div[1]/p")
            dataTelephone1 = inputElement.text
            if dataTelephone1 == 'Telefonnummer':
                dataTelephone1 = ""
            if dataTelephone1 == 'Telefonn...':
                dataTelephone1 = ""
        except:
            dataTelephone1 = ""
        f.write('"' + dataTelephone1 + '"' + ';')

        try:
            time.sleep(0.001)
            inputElement = driver.find_element_by_xpath("//span[text()='Telefonnummer, Fax, Mobil']/../div[1]/div[1]/div[1]/div[2]/p")
            dataTelephone2 = inputElement.text
            if dataTelephone2 == 'Faxnummer':
                dataTelephone2 = ""
            if dataTelephone2 == 'Faxnum...':
                dataTelephone2 = ""
        except:
            dataTelephone2 = ""
        f.write('"' + dataTelephone2 + '"' + ';')

        try:
            time.sleep(0.001)
            inputElement = driver.find_element_by_xpath("//span[text()='Telefonnummer, Fax, Mobil']/../div[1]/div[1]/div[1]/div[3]/p")
            dataTelephone3 = inputElement.text
            if dataTelephone3 == 'Mobilnummer':
                dataTelephone3 = ""
            if dataTelephone3 == 'Mobilnu...':
                dataTelephone3 = ""
        except:
            dataTelephone3 = ""
        f.write('"' + dataTelephone3 + '"' + ';')

        try:
            time.sleep(0.001)
            inputElement = driver.find_element_by_xpath("//span[text()='Telefonnummer, Fax, Mobil']/..//span[text()='zur Website']/..")
            href = inputElement.get_attribute("href")
        except:
            href = ""
        f.write('"' + href + '"')
        time.sleep(0.001)

        f.write('\n')

        time.sleep(0.001)

        if i == 0:
            driver.find_element_by_xpath(".//*[@id='entry-detail']/section/div/div[1]/div[2]/span/a").click()
        elif i < numberOfPage:
            try:
                inputElement = driver.find_element_by_xpath(".//*[@id='entry-detail']/section/div/div[1]/div[2]/span/a[2]").click()
            except:
                time.sleep(15)
                try:
                    inputElement = driver.find_element_by_xpath(
                        ".//*[@id='entry-detail']/section/div/div[1]/div[2]/span/a[2]").click()
                except:
                    time.sleep(15)
                    try:
                        inputElement = driver.find_element_by_xpath(
                            ".//*[@id='entry-detail']/section/div/div[1]/div[2]/span/a[2]").click()
                    except:
                        time.sleep(15)
                        try:
                            inputElement = driver.find_element_by_xpath(
                                ".//*[@id='entry-detail']/section/div/div[1]/div[2]/span/a[2]").click()
                        except:
                            time.sleep(15)
                            try:
                                inputElement = driver.find_element_by_xpath(
                                    ".//*[@id='entry-detail']/section/div/div[1]/div[2]/span/a[2]").click()
                            except:
                                time.sleep(15)
                                try:
                                    inputElement = driver.find_element_by_xpath(
                                        ".//*[@id='entry-detail']/section/div/div[1]/div[2]/span/a[2]").click()
                                except:
                                    time.sleep(15)
                                    try:
                                        inputElement = driver.find_element_by_xpath(
                                            ".//*[@id='entry-detail']/section/div/div[1]/div[2]/span/a[2]").click()
                                    except:
                                        time.sleep(15)
                                        try:
                                            inputElement = driver.find_element_by_xpath(
                                                ".//*[@id='entry-detail']/section/div/div[1]/div[2]/span/a[2]").click()
                                        except:
                                            time.sleep(15)
                                            try:
                                                inputElement = driver.find_element_by_xpath(
                                                    ".//*[@id='entry-detail']/section/div/div[1]/div[2]/span/a[2]").click()
                                            except:
                                                time.sleep(15)
                                                try:
                                                    inputElement = driver.find_element_by_xpath(
                                                        ".//*[@id='entry-detail']/section/div/div[1]/div[2]/span/a[2]").click()
                                                except:
                                                    time.sleep(15)
                                                    try:
                                                        inputElement = driver.find_element_by_xpath(
                                                            ".//*[@id='entry-detail']/section/div/div[1]/div[2]/span/a[2]").click()
                                                    except:
                                                        time.sleep(15)
                                                        try:
                                                            inputElement = driver.find_element_by_xpath(
                                                                ".//*[@id='entry-detail']/section/div/div[1]/div[2]/span/a[2]").click()
                                                        except:
                                                            time.sleep(15)
                                                            try:
                                                                inputElement = driver.find_element_by_xpath(
                                                                    ".//*[@id='entry-detail']/section/div/div[1]/div[2]/span/a[2]").click()
                                                            except:
                                                                time.sleep(15)
                                                                try:
                                                                    inputElement = driver.find_element_by_xpath(
                                                                        ".//*[@id='entry-detail']/section/div/div[1]/div[2]/span/a[2]").click()
                                                                except:
                                                                    time.sleep(15)
                                                                    try:
                                                                        inputElement = driver.find_element_by_xpath(
                                                                            ".//*[@id='entry-detail']/section/div/div[1]/div[2]/span/a[2]").click()
                                                                    except:
                                                                        time.sleep(15)
                                                                        try:
                                                                            inputElement = driver.find_element_by_xpath(
                                                                                ".//*[@id='entry-detail']/section/div/div[1]/div[2]/span/a[2]").click()
                                                                        except:
                                                                            time.sleep(15)
                                                                            print('End cycle')
                                                                        print('End cycle')
                                                                    print('End cycle')
                                                                print('End cycle')
                                                            print('End cycle')
                                                        print('End cycle')
                                                    print('End cycle')
                                                print('End cycle')
                                            print('End cycle')
                                        print('End cycle')
                                    print('End cycle')
                                print('End cycle')
                            print('End cycle')
                        print('End cycle')
                    print('End cycle')
                print('End cycle')
        else:
            break
        time.sleep(2)

    time.sleep(0.1)
    f.close()


def parsePageByXpath(catalogName, driver, city, number, xpath, path):

    time.sleep(5)
    print(city)
    driver.find_element_by_xpath(xpath + "/a[" + number + "]").click()

    time.sleep(5)
    driver.find_element_by_xpath(".//search-result-list/ol//li[1]//div/div[1]/div[1]/h2/a/span").click()
    time.sleep(5)

    time.sleep(0.6)
    inputElement = driver.find_element_by_xpath(".//*[@id='entry-detail']/section/div/div[1]/div[2]/span/span[3]")
    numberOfPage = int(inputElement.text)
    #numberOfPage = 10
    print(numberOfPage)
    f = open(path[0] + '/' + catalogName + '.csv', 'a', encoding="utf-8")

    for i in range(numberOfPage):
        if i % 20 == 0:
            # Obtain network adaptors configurations
            nic_configs = wmi.WMI().Win32_NetworkAdapterConfiguration(IPEnabled=True)
            # First network adaptor
            nic = nic_configs[0]
            # IP address, subnetmask and gateway values should be unicode objects
            ip = '192.168.0.' + str(random.randint(1, 99))
            subnetmask = '255.255.255.0'
            gateway = '192.168.0.' + str(random.randint(1, 99))
            # Set IP address, subnetmask and default gateway
            # Note: EnableStatic() and SetGateways() methods require *lists* of values to be passed
            nic.EnableStatic(IPAddress=[ip], SubnetMask=[subnetmask])
            nic.SetGateways(DefaultIPGateway=[gateway])
        print(i)
        pageNum = str(i+1)
        time.sleep(0.001)

        f.write('"' + catalogName + '"' + ';')

        try:
            inputElement = driver.find_element_by_xpath(".//*[@id='entry']/div/div[2]/div[1]/h1")
            name = inputElement.text
        except:
            name = ""
        f.write('"' + name + '"' + ';')

        time.sleep(0.001)
        try:
            inputElement = driver.find_element_by_xpath(".//*[@id='entry']/div/div[2]/div[3]/div[1]/p/span[1]/span")
            dataSpan1 = inputElement.text
            listOfNumbers = re.findall('\d+', dataSpan1)
            lenghtNumbers = len(listOfNumbers)
            lastNumber = listOfNumbers[lenghtNumbers - 1]
            startNumberOfHome = dataSpan1.find(lastNumber)
            street = dataSpan1[0:startNumberOfHome].lstrip(' ').rstrip(' ')
            numberOfHome = dataSpan1[startNumberOfHome:len(dataSpan1)].lstrip(' ').rstrip(' ')
            if lenghtNumbers > 1:
                preLastNumber = listOfNumbers[lenghtNumbers - 2]
                startPositionPreLastNumberOfHome = dataSpan1.find(preLastNumber)
                LastPositionOfPreLastNumberOfHome = startPositionPreLastNumberOfHome + len(str(preLastNumber))
                startNumberOfHome = dataSpan1.find(lastNumber)
                valueBetweenLastAndPreLastNumber = dataSpan1[LastPositionOfPreLastNumberOfHome:startNumberOfHome]
                if valueBetweenLastAndPreLastNumber.lstrip(' ').rstrip(' ') == '-':
                    street = dataSpan1[0:startPositionPreLastNumberOfHome].lstrip(' ').rstrip(' ')
                    numberOfHome = dataSpan1[startPositionPreLastNumberOfHome:len(dataSpan1)].lstrip(' ').rstrip(' ')
        except:
            street =""
            numberOfHome = ""
        f.write('"' + street + '"' + ';')
        f.write('"' + numberOfHome + '"' + ';')

        time.sleep(0.001)
        try:
            inputElement = driver.find_element_by_xpath(".//*[@id='entry']/div/div[2]/div[3]/div[1]/p/span[2]/span[1]")
            postcode = inputElement.text
        except:
            postcode = ""
        f.write('"' + postcode + '"' + ';')

        time.sleep(0.001)
        try:
            inputElement = driver.find_element_by_xpath(".//*[@id='entry']/div/div[2]/div[3]/div[1]/p/span[2]")
            dataSpan2 = inputElement.text
            city = dataSpan2[len(postcode):].lstrip(' ').rstrip(' ')
        except:
            city = ""
        f.write('"' + city + '"' + ';')
        f.write('"Deutschland"' + ';')

        time.sleep(0.001)
        lastRightUrl = driver.current_url
        try:
            driver.find_element_by_xpath("//span[text()='Telefonnummer, Fax, Mobil']/..//span[text()='Kontakt anzeigen!']/..").click()
        except:
            print()

        try:
            time.sleep(1)
            inputElement = driver.find_element_by_xpath("//span[text()='Telefonnummer, Fax, Mobil']/../div[1]/div[1]/div[1]/div[1]/p")
            dataTelephone1 = inputElement.text
            if dataTelephone1 == 'Telefonnummer':
                dataTelephone1 = ""
            if dataTelephone1 == 'Telefonn...':
                dataTelephone1 = ""
        except:
            dataTelephone1 = ""
        f.write('"' + dataTelephone1 + '"' + ';')

        try:
            time.sleep(0.001)
            inputElement = driver.find_element_by_xpath("//span[text()='Telefonnummer, Fax, Mobil']/../div[1]/div[1]/div[1]/div[2]/p")
            dataTelephone2 = inputElement.text
            if dataTelephone2 == 'Faxnummer':
                dataTelephone2 = ""
            if dataTelephone2 == 'Faxnum...':
                dataTelephone2 = ""
        except:
            dataTelephone2 = ""
        f.write('"' + dataTelephone2 + '"' + ';')

        try:
            time.sleep(0.001)
            inputElement = driver.find_element_by_xpath("//span[text()='Telefonnummer, Fax, Mobil']/../div[1]/div[1]/div[1]/div[3]/p")
            dataTelephone3 = inputElement.text
            if dataTelephone3 == 'Mobilnummer':
                dataTelephone3 = ""
            if dataTelephone3 == 'Mobilnu...':
                dataTelephone3 = ""
        except:
            dataTelephone3 = ""
        f.write('"' + dataTelephone3 + '"' + ';')

        try:
            time.sleep(0.001)
            inputElement = driver.find_element_by_xpath(
                "//span[text()='Telefonnummer, Fax, Mobil']/..//span[text()='zur Website']/..")
            href = inputElement.get_attribute("href")
        except:
            href = ""
        f.write('"' + href + '"')
        time.sleep(0.001)

        f.write('\n')

        time.sleep(0.001)

        if i == 0:
            driver.find_element_by_xpath(".//*[@id='entry-detail']/section/div/div[1]/div[2]/span/a").click()
        elif i < numberOfPage:
            try:
                inputElement = driver.find_element_by_xpath(".//*[@id='entry-detail']/section/div/div[1]/div[2]/span/a[2]").click()
            except:
                print('End cycle')
        else:
            break
        time.sleep(0.2)

    time.sleep(0.1)
    f.close()

def writeHeader(catalogName, path):
    f = open(path[0] + '/' + catalogName + '.csv', 'a', encoding="utf-8")
    f.write("\"branch\";\"company-name\";\"street\";\"house\";\"postcode\";\"city\";\"country\";\"phone\";\"fax\";\"mobile\";\"website\"")
    f.write('\n')
    f.close()