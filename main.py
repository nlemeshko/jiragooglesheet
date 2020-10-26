import requests
from pandas import DataFrame, read_csv
import pandas as pd
import gspread
import time
from datetime import datetime

tasks = list()
times = list()
task = list()
task2 = list()

t = datetime.now()
m=t.strftime("%b")
d=t.strftime("%d")



url='/secure/ConfigureReport!excelView.jspa?htmlExport=true&startDate=1%2FFeb%2F20&endDate='+ d + '%2F' + m + '%2F20&reportKey=jira-timesheet-plugin:report&projectid=13402&weekends=&showDetails=&monthView=true&sum=week&reportingDay=1'
headers = {'Content-Type': 'application/json','Authorization': 'Basic '}
r = requests.get(url, headers=headers, stream=True)

dfs = pd.read_html(r.content)
df = pd.DataFrame(dfs[0])
for i in range(len(df)):
    tasks.append(df.loc[i][2])
    times.append(df.loc[i][5])

tasks.pop(0)
tasks.pop(0)
tasks.pop(-1)

times.pop(0)
times.pop(0)
times.pop(-1)

newdf = pd.DataFrame(list(zip(tasks, times)),
               columns =['Tasks', 'Time'])

gc = gspread.service_account(filename='client_secret.json')
wks1 = gc.open("Miraworks 2.0 - Замечания")
wks=wks1.worksheet("Сдача МВП 2.0")
list_of_hashes = wks.col_values(6)


for i in range(len(list_of_hashes)):
    task.append(list_of_hashes[i][-9:])


for i in range(len(list_of_hashes)):
    if 'h' in newdf.loc[newdf['Tasks'] == task[i]]['Time'].to_string(index=False):
       if i == 90:
           time.sleep(100)
       wks.update_cell(i+1, 7, newdf.loc[newdf['Tasks'] == task[i]]['Time'].to_string(index=False))

time.sleep(100)


wks=wks1.worksheet("Тех. Поддержка")
list_of_hashes = wks.col_values(8)

for i in range(len(list_of_hashes)):
    task2.append(list_of_hashes[i][-9:])


for i in range(len(list_of_hashes)):
    if 'h' in newdf.loc[newdf['Tasks'] == task2[i]]['Time'].to_string(index=False):
      if i == 90:
            time.sleep(100)
      wks.update_cell(i+1, 9, newdf.loc[newdf['Tasks'] == task2[i]]['Time'].to_string(index=False))