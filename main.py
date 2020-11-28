import requests
import pandas as pd
import gspread
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime


tasks = list()
times = list()
newtime = list()
newtime2 = list()
newbranches = list()
newbranches2 = list()
branches = list()
task = list()
task2 = list()

t = datetime.now()
m=t.strftime("%b")
d=t.strftime("%d")

jira = '/secure/ConfigureReport!excelView.jspa?htmlExport=true&startDate=1%2FFeb%2F20&endDate='+ d + '%2F' + m + '%2F20&reportKey=jira-timesheet-plugin:report&projectid=13402&weekends=&showDetails=&monthView=true&sum=week&reportingDay=1'
print('Starting download jira excel...')
headers = {'Content-Type': 'application/json', 'Authorization': 'Basic '}
r = requests.get(jira, headers=headers, stream=True)

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


headers2 = {'PRIVATE-TOKEN': ''}
gitlabback = '/api/v4/projects/137/search?scope=commits&search='
gitlabfront = '/api/v4/projects/161/search?scope=commits&search='
gitlabadmin = '/api/v4/projects/137/search?scope=commits&search='

branches=tasks.copy()

def backendsearch(i):
        taskindex=branches.index(i)
        res = requests.get(gitlabback + i + '&ref=production', headers=headers2)
        response = res.json()
        if response:
            branches.pop(taskindex)
            branches.insert(taskindex,'production')
        else:
            res = requests.get(gitlabback + i + '&ref=master', headers=headers2)
            response = res.json()
            if response:
                branches.pop(taskindex)
                branches.insert(taskindex, 'master')
            else:
                res = requests.get(gitlabback + i + '&ref=test', headers=headers2)
                response = res.json()
                if response:
                    branches.pop(taskindex)
                    branches.insert(taskindex, 'test')
                else:
                    res = requests.get(gitlabback + i + '&ref=develop', headers=headers2)
                    response = res.json()
                    if response:
                        branches.pop(taskindex)
                        branches.insert(taskindex, 'develop')
        return print('Done '+i)

def frontsearch(i):
        taskindex=branches.index(i)
        if 'MIRA' in i:
            res = requests.get(gitlabfront + i + '&ref=production', headers=headers2)
            response = res.json()
            if response:
                branches.pop(taskindex)
                branches.insert(taskindex, 'production')
            else:
                res = requests.get(gitlabfront + i + '&ref=master', headers=headers2)
                response = res.json()
                if response:
                    branches.pop(taskindex)
                    branches.insert(taskindex, 'production')
                else:
                    res = requests.get(gitlabfront + i + '&ref=test', headers=headers2)
                    response = res.json()
                    if response:
                        branches.pop(taskindex)
                        branches.insert(taskindex, 'production')
                    else:
                        res = requests.get(gitlabfront + backend[item] + '&ref=develop', headers=headers2)
                        response = res.json()
                        if response:
                            branches.pop(taskindex)
                            branches.insert(taskindex, 'production')
        return print('Done ' +i)


def adminsearch(i):
    taskindex = branches.index(i)
    if 'MIRA' in i:
        res = requests.get(gitlabadmin + i + '&ref=production', headers=headers2)
        response = res.json()
        if response:
            branches.pop(taskindex)
            branches.insert(taskindex, 'production')
        else:
            res = requests.get(gitlabadmin + i + '&ref=master', headers=headers2)
            response = res.json()
            if response:
                branches.pop(taskindex)
                branches.insert(taskindex, 'production')
            else:
                res = requests.get(gitlabadmin + i + '&ref=test', headers=headers2)
                response = res.json()
                if response:
                    branches.pop(taskindex)
                    branches.insert(taskindex, 'production')
                else:
                    res = requests.get(gitlabadmin + backend[item] + '&ref=develop', headers=headers2)
                    response = res.json()
                    if response:
                        branches.pop(taskindex)
                        branches.insert(taskindex, 'production')
    return print('Done ' + i)


print('Starting parsing Backend...')
with ThreadPoolExecutor() as executor:
    executor.map(backendsearch, branches)
print('End parsing Backend.')

print('Starting parsing Frontend...')
with ThreadPoolExecutor() as executor:
    executor.map(frontsearch, branches)
print('End parsing Frontend.')

print('Starting parsing Admin...')
with ThreadPoolExecutor() as executor:
    executor.map(adminsearch, branches)
print('End parsing Admin.')



newdf = pd.DataFrame(list(zip(tasks, times, branches)),
                         columns=['Tasks', 'Time', 'Branches'])



gc = gspread.service_account(filename='client_secret.json')
wks1 = gc.open("Miraworks 2.0 - Замечания")
wks=wks1.worksheet("Сдача МВП 2.0")
list_of_hashes = wks.col_values(6)




for i in range(len(list_of_hashes)):
    task.append(list_of_hashes[i][-9:])



for i in range(len(task)):
    if task[i] in tasks:
        newtime.append(newdf.loc[newdf['Tasks'] == task[i]]['Time'].to_string(index=False))
    else:
        newtime.append('')

for i in range(len(task)):
    if task[i] in tasks:
        newbranches.append(newdf.loc[newdf['Tasks'] == task[i]]['Branches'].to_string(index=False))
    else:
        newbranches.append('')


lastdf = pd.DataFrame(list(zip(task, newtime, newbranches)),
                          columns=['Task', 'Time', 'Branches'])

x=0
range1 = wks.range('G1:G'+str(lastdf.index[-1]))
for cell in range1:
    cell.value = lastdf['Time'][x]
    x=x+1

wks.update_cells(range1)

x=0
range1 = wks.range('H1:H'+str(lastdf.index[-1]))
for cell in range1:
    cell.value = lastdf['Branches'][x]
    x=x+1

wks.update_cells(range1)



wks2=wks1.worksheet("Тех. поддержка")
list_of_hashes2 = wks2.col_values(8)

for i in range(len(list_of_hashes2)):
    task2.append(list_of_hashes2[i][-9:])



for i in range(len(task2)):
    if task2[i] in tasks:
        newtime2.append(newdf.loc[newdf['Tasks'] == task2[i]]['Time'].to_string(index=False))
    else:
        newtime2.append('')

for i in range(len(task2)):
    if task2[i] in tasks:
        newbranches2.append(newdf.loc[newdf['Tasks'] == task2[i]]['Branches'].to_string(index=False))
    else:
        newbranches2.append('')


lastdf2 = pd.DataFrame(list(zip(task2, newtime2, newbranches2)),
                          columns=['Task', 'Time', 'Branches'])


x=0
range1 = wks2.range('I1:I'+str(lastdf2.index[-1]))
for cell in range1:
    cell.value = lastdf2['Time'][x]
    x=x+1

wks2.update_cells(range1)

x=0
range1 = wks2.range('J1:J'+str(lastdf2.index[-1]))
for cell in range1:
    cell.value = lastdf2['Branches'][x]
    x=x+1

wks2.update_cells(range1)

wks2.update('I9','Time')
wks2.update('J9','Branches')
