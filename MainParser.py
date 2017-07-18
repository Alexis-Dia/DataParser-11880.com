from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import time
import csv
import sys
import re
import os
import wmi
import random
from PageParser import parsePage, parsePageWithCatalogId, parsePageSlow, parsePageByXpath, writeHeader

phantomjs_path = "D:\phantomjs.exe"
driver = webdriver.PhantomJS(executable_path=phantomjs_path, service_log_path=os.path.devnull)

cities = ['Berlin', 'Hamburg', 'München', 'Köln', 'Frankfurt', 'Düsseldorf', 'Stuttgart', 'Bremen', 'Essen', 'Nürnberg',
          'Dortmund', 'Hannover', 'Duisburg',
          'Dresden', 'Leipzig', 'Wuppertal', 'Bielefeld', 'Bochum', 'Bonn', 'Münster', 'Karlsruhe', 'Wiesbaden',
          'Mannheim', 'Augsburg', 'Braunschweig',
          'Mönchengladbach', 'Aachen', 'Kiel', 'Krefeld', 'Freiburg', 'Lübeck', 'Gelsenkirchen', 'Mainz', 'Kassel',
          'Saarbrücken', 'Chemnitz', 'Oldenburg',
          'Osnabrück', 'Hagen', 'Oberhausen', 'Hamm', 'Solingen', 'Magdeburg', 'Mülheim', 'Darmstadt', 'Halle',
          'Erfurt', 'Neuss', 'Regensburg', 'Leverkusen',
          'Paderborn', 'Heidelberg', 'Würzburg', 'Ludwigshafen', 'Rostock', 'Herne', 'Ingolstadt', 'Ulm',
          'Bergisch Gladbach', 'Koblenz', 'Remscheid',
          'Reutlingen', 'Heilbronn', 'Göttingen', 'Siegen', 'Potsdam', 'Fürth', 'Pforzheim', 'Erlangen',
          'Recklinghausen', 'Wolfsburg', 'Trier', 'Gütersloh',
          'Bottrop', 'Salzgitter', 'Iserlohn', 'Kaiserslautern', 'Bremerhaven', 'Moers', 'Hildesheim', 'Offenbach',
          'Ratingen', 'Konstanz', 'Witten',
          'Norderstedt', 'Esslingen', 'Villingen-Schwenningen', 'Wilhelmshaven', 'Tübingen', 'Rheine', 'Velbert',
          'Minden', 'Bocholt', 'Arnsberg',
          'Jena', 'Ludwigsburg', 'Detmold', 'Hanau', 'Düren', 'Dorsten']

path = ['D:/parseData/']

writeHeader('Elektriker', path)
url = "https://www.11880-elektriker.com/elektriker"
for city in cities:
    driver.get(url)
    parsePage('Elektriker', driver, city, path)

writeHeader('Werkstatt', path)
url = "https://www.11880-werkstatt.com/werkstatt"
for city in cities:
    time.sleep(5)
    driver.get(url)
    writeHeader('Werkstatt', path)
    parsePage('Werkstatt', driver, city, path)

writeHeader('Dachdecker', path)
url = "https://www.11880-dachdecker.com/dachdecker"
for city in cities:
    time.sleep(5)
    driver.get(url)
    parsePage('Dachdecker', driver, city, path)

writeHeader('Maler', path)
url = "https://www.11880-maler.com/maler"
for city in cities:
    time.sleep(5)
    driver.get(url)
    parsePage('Maler', driver, city, path)

writeHeader('Heizung', path)
url = "https://www.11880-heizung.com/heizung"
for city in cities:
    time.sleep(5)
    driver.get(url)
    parsePage('Heizung', driver, city, path)

writeHeader('Tischler', path)
url = "https://www.11880-tischler.com/tischler"
for city in cities:
    time.sleep(5)
    driver.get(url)
    parsePage('Tischler', driver, city, path)

writeHeader('Gartenbau', path)
url = "https://www.11880-gartenbau.com/gartenbau"
for city in cities:
    time.sleep(5)
    driver.get(url)
    parsePage('Gartenbau', driver, city, path)

writeHeader('Umzug', path)
url = "https://www.11880-umzug.com/umzug"
for city in cities:
    time.sleep(5)
    driver.get(url)
    parsePage('Umzug', driver, city, path)

writeHeader('Gebäudereiniger', path)
url = "https://www.11880-gebaeudereinigung.com/gebaeudereinigung"
for i in enumerate(cities):
    time.sleep(5)
    driver.get(url)
    position = str(i[0] + 1)
    city = i[1]
    xpath = ".//div/div/section/div/div/div/div"
    parsePageByXpath('Gebäudereiniger', driver, city, position, xpath, path)

writeHeader('Immobilienmakler', path)
url = "https://www.11880-immobilienmakler.com/immobilienmakler"
for city in cities:
    time.sleep(5)
    driver.get(url)
    parsePage('Immobilienmakler', driver, city, path)

writeHeader('Rechtsanwalt', path)
url = "https://www.11880-rechtsanwalt.com/rechtsanwalt"
for city in cities:
    time.sleep(5)
    driver.get(url)
    parsePage('Rechtsanwalt', driver, city, path)

writeHeader('Zahnarzt', path)
url = "https://www.11880-zahnarzt.com/zahnarzt"
for city in cities:
    time.sleep(5)
    driver.get(url)
    parsePage('Zahnarzt', driver, city, path)

writeHeader('Steuerberater', path)
url = "https://www.11880-steuerberater.com/steuerberater"
for city in cities:
    time.sleep(5)
    driver.get(url)
    parsePage('Steuerberater', driver, city, path)

writeHeader('Versicherung', path)
url = "https://www.11880-versicherung.com/versicherung"
for city in cities:
    time.sleep(5)
    driver.get(url)
    parsePage('Versicherung', driver, city, path)

writeHeader('Bestatter', path)
url = "https://www.11880-bestattung.com/bestattung"
for i in enumerate(cities):
    time.sleep(5)
    driver.get(url)
    position = str(i[0] + 1)
    city = i[1]
    parsePageWithCatalogId('Bestatter', driver, city, position, path)

writeHeader('Friseur', path)
url = "https://www.11880-beauty.com/friseur"
for i in enumerate(cities):
    time.sleep(5)
    driver.get(url)
    position = str(i[0] + 1)
    city = i[1]
    parsePageWithCatalogId('Friseur', driver, city, position, path)

writeHeader('Physiotherapeut', path)
url = "https://www.11880-physio.com/physio"
for i in enumerate(cities):
    time.sleep(5)
    driver.get(url)
    position = str(i[0] + 1)
    city = i[1]
    parsePageWithCatalogId('Physiotherapeut', driver, city, position, path)

writeHeader('Nagelstudio', path)
url = "https://www.11880-beauty.com/nagelstudio"
for i in enumerate(cities):
    time.sleep(5)
    driver.get(url)
    position = str(i[0] + 1)
    city = i[1]
    parsePageSlow('Nagelstudio', driver, city, position, path)

writeHeader('Kosmetikstudio', path)
url = "https://www.11880-beauty.com/kosmetikstudio"
for i in enumerate(cities):
    time.sleep(5)
    driver.get(url)
    position = str(i[0] + 1)
    city = i[1]
    parsePageSlow('Kosmetikstudio', driver, city, position, path)

x = driver.page_source
driver.close()