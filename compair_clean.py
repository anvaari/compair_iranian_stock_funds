#Import Packages
import pandas as pd
import numpy as np
import os
from persiantools.jdatetime import JalaliDate as dt


#Inputs
exls_dir=r'' #Directory where excels file exist
shakhes_name='' #Name of xlsx file contain shakhes value
month='' #Which month do you want to start from ---> in this form (str) : '05' or '10'
day_range=(,) #Range of days in month _ For example: (30,10) it start from 30th day of month and go back (to avoid errors)

'''
This function get NAVs and return dictionary where its keys are years of stock fund and each value is revenue during that year. 
In mode 1 function calculate it year by year
In mode 2 fuunction calculate it every 2 years
In mode 5 fuunction calculate it every 5 years

'''
def bazdeh_calc(data,data_ebtal,mode):
    bazdeh={}
    keys=list(data.keys())
    if mode==1:
        for year in range(keys[0],keys[-1]):
            bazdeh[year]=(data_ebtal[year+1]-data[year])/data[year]*100
    elif mode==2:
        if len(data)<2: #check data have at least data of 2 year 
            return np.nan
        for year in range(keys[0],keys[-2]):
            bazdeh[year]=(data_ebtal[year+2]-data[year])/data[year]*100
    elif mode==5:
        if len(data)<6: #check data have at least data of 5 year 
            return np.nan
        else:
            bazdeh[keys[-6]]=(data_ebtal[keys[-1]]-data[keys[-6]])/data[keys[-6]]*100
            
    return bazdeh
            

names=os.listdir(exls_dir) #Get list of file in exls_dir directory
names.remove(shakhes_name) 
result=pd.DataFrame(columns=['Name','Sal Tasis']) #Creat final dataframe. now it is empty
names=['/'+x.strip('\n').strip('~$') for x in names]
shakhes=pd.read_excel(exls_dir+'/'+shakhes_name) #Read Shakhes data from its file

shakhes_data=pd.DataFrame()
for name in names: #Loop on every stock fund's name
    nav_yearly={} 
    nav_yearly_ebtal={}
    temp = pd.read_excel(exls_dir+name) #This dataframe hold data of "name" in each loop.
    afz=False #Boolean for detect Capital Increase (Afzayesh Sarmaye)
    date_afz=dt(2000,2,2) #Date of Capital Increase (Afzayesh Sarmaye)
    for index, row in temp.iterrows(): #This loop itrate over every item and try to detect if Capital Increase (Afzayesh Sarmaye) exist or not.
        try:
            if row['NAVSodoor']/temp.loc[index+1,'NAVSodoor'] <= 0.5:
                afz=True
                temp_date=temp.loc[index,'Date']
                date_afz=dt(int(temp_date[:4]),int(temp_date[5:7]),int(temp_date[-2:]))
                factor=temp.loc[index+1,'NAVSodoor']/row['NAVSodoor'] #This factor used for balance NAVs after Capital Increase (Afzayesh Sarmaye)
                factor_ebtal=temp.loc[index+1,'NAVEbtal']/row['NAVEbtal']
        except :
            pass
    date_start=temp.iloc[-1]['Date'] #Start date of activity of "name"
    year_start=int(date_start[:4])
    year_end=int(temp.iloc[0]['Date'][:4])
    for j in range(year_start,year_end+1): #in this loop we creat Dictionary contain NAV of each year in specified date.
        if int(date_start[5:7])>int(month)+1 and j==int(date_start[:4]): #Check if month of start greater than month we wanth to analyze else go to next year
            year_start+=1
            continue
        for i in range(day_range[0],day_range[1],-1):
            date_now='{}{}{}'.format(j,f'/{month}/',i)
            date_now_dt=dt(int(date_now[:4]),int(date_now[5:7]),int(date_now[-2:]))
            if j==1396 and name=='/Melli.xlsx': #Melli fund doesn't have data for month 5 of year 1396
                date_now='{}{}{}'.format(j,'/06/',i)
            nav_now=temp[temp['Date']==date_now]['NAVSodoor']
            nav_now_ebtal=temp[temp['Date']==date_now]['NAVEbtal']
            if date_now_dt>date_afz:
                nav_now=nav_now*factor
                nav_now_ebtal=nav_now_ebtal*factor_ebtal
            if len(nav_now)==0:
                continue
            else:
                nav_yearly[j]=nav_now.iloc[0]
                nav_yearly_ebtal[j]=nav_now_ebtal.iloc[0]
                break
            
    bazdeh1=bazdeh_calc(nav_yearly,nav_yearly_ebtal, 1)
    bazdeh2=bazdeh_calc(nav_yearly,nav_yearly_ebtal, 2)
    bazdeh5=bazdeh_calc(nav_yearly,nav_yearly_ebtal, 5)
    
    #Put data on final DataFrame
    for year in bazdeh1.keys(): #Put revenues on DataFrame
        result.at[names.index(name),year]=bazdeh1[year]
        
    result.at[names.index(name),'Name']=temp.loc[0]['name']
    result.at[names.index(name),'Sal Tasis']=int(date_start[:4])
    result.at[names.index(name),'NAV Sodoor']=temp.loc[0]['NAVSodoor']
    result.at[names.index(name),'Miangin Salane']=pd.Series(bazdeh1).mean()
    result.at[names.index(name),'Miangin 2 sal']=pd.Series(bazdeh2).mean()
    result.at[names.index(name),'5 sal']=pd.Series(bazdeh5).iloc[0]
    
    #Calculate total revenue of "name" 
    nav_first=temp.iloc[-1]['NAVSodoor']
    nav_last=temp.iloc[0]['NAVEbtal']
    if afz==True:
        nav_last=nav_last*factor_ebtal
    result.at[names.index(name),'Kol']=(nav_last-nav_first)/nav_first*100

#Calculate revenue of 'Shakhes kol'
shakhes_yearly={}
for j in range(1389,1400):
    for i in range(day_range[0],day_range[1],-1):
        try:
            date_now_shakhes=int('{}{}{}'.format(j,month,i))
            shakhes_now=shakhes[shakhes['dateissue']==date_now_shakhes]['Value']
            shakhes_yearly[j]=shakhes_now.iloc[0]
        except:
            continue
        else:
            break
            
bazdeh_shakhes1=bazdeh_calc(shakhes_yearly, shakhes_yearly, 1)
bazdeh_shakhes2=bazdeh_calc(shakhes_yearly, shakhes_yearly, 2)
bazdeh_shakhes5=bazdeh_calc(shakhes_yearly, shakhes_yearly, 5)
    
for year in bazdeh_shakhes1.keys():
    shakhes_data.at[0,year]=bazdeh_shakhes1[year]
shakhes_data.at[0,'Miangin Salane']=pd.Series(bazdeh_shakhes1).mean()
shakhes_data.at[0,'Miangin 2 sal']=pd.Series(bazdeh_shakhes2).mean()
shakhes_data.at[0,'5 sal']=pd.Series(bazdeh_shakhes5).iloc[0]
shakhes_data.at[0,'Kol']=(shakhes.iloc[0]['Value']-shakhes.iloc[-1]['Value'])/shakhes.iloc[-1]['Value']*100


#Calculate mean score (if revenue of stock fund greater than mean of others in that year score +1) and Shakhes Score (if revenue of investment fund greater than revenue of 'Shakhes' in that year score +1) in each year
year_info=pd.DataFrame(result.loc[:,[i for i in range(1389,1399)]].describe()) #Summary of result DataFrame
for index,row in result.iterrows():
    score_m=0
    score_sh=0
    for i in range(1389,1400):
        try:
            if row[i]>year_info.at['mean',i]:
                score_m+=1
            if row[i]>shakhes_data.at[0,i]:
                score_sh+=1
        except:
            continue

    result.at[index,'Score Miangin']=score_m
    result.at[index,'Score Shakhes']=score_sh
    
col=[i for i in range(1389,1399)]
col.append('Score Miangin')
col.append('Score Shakhes')
year_info=pd.DataFrame(result.loc[:,col].describe())



#Sort result in different ways
best_kol=pd.DataFrame(result.sort_values(by=['Kol'],ascending=False)).reset_index(drop=True)
best_5=pd.DataFrame(result.sort_values(by=['5 sal'],ascending=False)).reset_index(drop=True)
best_mian=pd.DataFrame(result.sort_values(by=['Score Miangin'],ascending=False)).reset_index(drop=True)
best_shakhes=pd.DataFrame(result.sort_values(by=['Score Shakhes'],ascending=False)).reset_index(drop=True)

#Select Best stock fund
for name in best_kol.head(10)['Name'].values:
    if name in best_5.head(5)['Name'].values and name in best_mian.head(12)['Name'].values and name in best_shakhes.head(10)['Name'].values:
        print(name)



result.to_excel('Final.xlsx')
year_info.to_excel('Final_Summary.xlsx')
