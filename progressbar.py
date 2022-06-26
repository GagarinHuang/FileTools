import PySimpleGUI as sg
import pandas as pd
import time
import openpyxl
from tqdm import tqdm
from itertools import islice

path = "C:\\Users\\Lenovo\Desktop\\make money\\拆分\\test\\汇总2.xlsx"
df = pd.read_excel(path)
'''
layout = [[sg.Text('任务完成进度')],
          [sg.Text('', size=(5, 1), font=('Helvetica', 15), justification='center', key='text')],
          [sg.ProgressBar(len(df), orientation='h', size=(50, 20), key='progressbar')],
          [sg.Cancel()]]
window = sg.Window('机器人执行进度', layout)
progress_bar = window['progressbar']
# For循环
for i in range(0, len(df)):
    print(df.iloc[i])
    time.sleep(1) #假设处理的时间
    event, values = window.read(timeout=10)
    if event == 'Cancel' or event is None:
        break
    progress_bar.UpdateBar(i + 1)
    window['text'].update('{}%'.format(int(i / len(df) * 100)))
window.close()
'''

for i in tqdm(range(100)):
    wb = openpyxl.load_workbook(path)
    ws = wb["个税汇总"]
    data = ws.values
    cols = next(data)[0:]
    data = list(data)
    data = (islice(r, 0, None) for r in data)
    df = pd.DataFrame(data, columns=cols)
    if not df.empty:
        break
    pass