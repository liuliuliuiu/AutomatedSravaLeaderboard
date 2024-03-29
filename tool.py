from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

import schedule
import time

import re
import openpyxl


PATH = "C:\Program Files (x86)\chromedriver.exe"
driver = webdriver.Chrome(PATH)

driver.get("https://www.strava.com/clubs/marianopolis")
print(driver.title)

stats = []
zzy = driver.find_elements(By.CLASS_NAME, "leaderboard")

for info in zzy:
    stats.append(info.text)

athletes = []
for item in stats:
    
    stats = item.split('\n')
stats.remove(stats[0])
for element in stats:
    if len(element) < 2:
        stats.remove(element)

BigStats = []
StatsLists = []
for element in stats:
    BigStats = element.split(" ")
    StatsLists.append(BigStats)


headers = ["Name", "", "Distance"]

athletestats = []

for athlete_stats in StatsLists:
    name = athlete_stats[0] + athlete_stats[1]
    athlete_dict = dict(zip(headers, [name, athlete_stats[2]]))
    athletestats.append(athlete_dict)


athletestats = [{'Name': athlete['Name'], 'Distance': athlete['']}
                for athlete in athletestats]
print(athletestats)


for i in StatsLists:
    name = i[0] + " " + i[1]
    i[0] = name
    i.pop(1)
print(StatsLists)

StatsDict = {}
for sublist in StatsLists:
    name = sublist[0]
    values = sublist[1]
    StatsDict[name] = values
print(StatsDict)
filename = r"C:\Users\liuli\OneDrive\Documents\SupposedlyAutomatedStravaLeaderboard.xlsx"
workbook = openpyxl.load_workbook(filename)
worksheet = workbook["Sheet1"]


for cell in worksheet[5]:
    if cell.value is None:
        column_letter = cell.column_letter
        break

keys = StatsDict.keys()
new_dict = {key: key.split()[-1].strip('.') for key in StatsDict}

for name, distance in StatsDict.items():
    if "Owen L." in name:
        worksheet[column_letter + "5"].value = float(distance)
        print(column_letter)


for row in range(6, worksheet.max_row + 1):
    cell_value = str(worksheet["B" + str(row)].value)
    if cell_value == "None":
        break
    fullname = (str(cell_value))
    print(fullname)
    last_name, first_name = fullname.split(', ')
    
    last_initial = last_name[0]
    
    updated_name = f"{first_name} {last_initial}."
    print(updated_name)
    if (updated_name) in StatsDict:
        worksheet[column_letter +
                  str(row)].value = float(StatsDict[updated_name])
workbook.save(filename)
print(StatsDict)
