
import pandas as pd
import os,re
def get_excelname(path):                         #get all excel name into a list
    excelname = []
    for file in os.listdir(path):
        if os.path.splitext(file)[1] == '.xlsx':
            excelname.append(file)
    return excelname



def findAllFile(path):
    for root, dirs, files in os.walk(path):
        for file in files:
            if re.match(r'Database.*(\.xlsx)$', file):
                fullname = os.path.join(root, file)
                yield fullname

#merge all files togrther
df = pd.DataFrame()
for i in findAllFile('.'):
    
    df=df.append(pd.read_excel(i,engine='openpyxl'),ignore_index=True)

df3 = pd.read_excel('放行总表.xlsx')
df=pd.merge(df,df3[['合并','Reports No.','最新状态']],how='left',on='合并')
df.to_excel('mergedfile.xlsx')

df.drop(df.columns[[0,1,2,3,4,5,7,13,15,17,18,20,21,22,24,25,26,27,28,29,30,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64]],axis=1,inplace=True)

a=input('please submit date...')
alist=a.split(',')
alist

df=df[df['Visual Date'].isin(alist)]
df1=df[df['RT Percent (%)']==100]

grouped = df.groupby(['RT Percent (%)','System Code','WPS NO','Size']).apply(lambda x:x.sample(frac=0.3)).reset_index(drop=True)
grouped1 = df.groupby(['RT Percent (%)','System Code','WPS NO','Size']).apply(lambda x:x.sample(frac=1))

grouped.to_excel('RTsample.xlsx')

df2=pd.DataFrame(grouped1)
df2.to_excel('Groupby.xlsx')

df1.to_excel('RTsampleBy100.xlsx')

