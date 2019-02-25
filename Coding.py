import pandas as pd
import datetime
import xlrd
d = pd.ExcelFile('\AAnm.xlsx')
df = d.parse('sheet1', skiprows = 34)
del df['Storage charge as follows']
for a in range(df.shape[0]):
    if not isinstance(df['Unnamed: 1'][a], basestring):
        df=df.drop([a])
        
b = 0 
while b < df.shape[0] and df['Unnamed: 3'][b].month != 1:
    b +=1
print df['Unnamed: 3'][b]

df1 = df[:b]
df1 = df1.sort_values(by = ['Unnamed: 6','Unnamed: 11'])

df2 = df[b:]
df2 = df2.sort_values(by = ['Unnamed: 6','Unnamed: 11'])

df3 = pd.concat([df1,df2])
df3 = df3.reset_index()
del df3['index']

WEEK = []
for r in range(df3.shape[0]):
    wee = df3['Unnamed: 6'][r].split('-')[0]
    da = wee.split('/')[0]
    mo = wee.split('/')[1]
    ye = wee.split('/')[2]
    we = pd.Timestamp(year=int(ye), month=int(mo), day=int(da))
    we = we.week
    WEEK.append(we)
 
df3['WEEK'] = WEEK

max_add_row = 3
i = 0
while i < df3.shape[0]:
    if df3['Unnamed: 11'][i] == 1.44:
        df3['Unnamed: 12'][i] = 1
    if df3['Unnamed: 11'][i] > 1.44:
        df3['Unnamed: 12'][i] = int(df3['Unnamed: 11'][i]/1.44)+1
    if df3['Unnamed: 11'][i] < 1.44:
        start = i
        y = 1
        sum_sku = df3['Unnamed: 11'][start]
        while start+1<df3.shape[0] and y < max_add_row and sum_sku + df3['Unnamed: 11'][start+1] < 1.44 and df3['WEEK'][start] == df3['WEEK'][start+1]:
            start+= 1
            sum_sku += df3['Unnamed: 11'][start]
            y += 1
            
        df3['Unnamed: 12'][start] = int(sum_sku/1.44)+1
        i = start
    i += 1
    
MON = []
s = pd.date_range('2019-01-01', '2019-01-31', freq='D').to_series()
for t in s:
    if t.weekday() == 0:
        MON.append(t)
        
df3.to_csv('tt.csv')
