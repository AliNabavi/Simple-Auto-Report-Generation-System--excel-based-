# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

#%% IMPORT LIBRARIES
import path,pandas,numpy,matplotlib,os,re,pathlib,jdatetime, xlrd ,jpype, asposecells , jdatetime
import func_var 



#%% SETTINGS
pandas.options.display.float_format = '{:.2f}'.format
directory = func_var.directory

#%% PARAMETERS

total_report_df = pandas.DataFrame(columns=func_var.total_report_df_columns)
monthes_to_name = func_var.monthes_to_name
        

#%% MONTH OPERATIN

total_billings = []

driver_dates_total_dic = {}
today_date = jdatetime.date.today()

for month in func_var.GET_MONTHES_LIST():    
        
    month_billings = []
    driver_dates_month_dic = {}
    
    month_folder = directory + "\\" + month
    
    for file in os.listdir(month_folder):    
        if file[:3].isdigit():
            total_billings.append(month_folder + "\\" + file)
            month_billings.append(month_folder + "\\" + file)
            # print(month_folder + "\\" + file)
   
    month_df_total = pandas.read_excel(month_billings[0],header=1)
    driver = month_df_total.columns[0]
    date = month_df_total.columns[-1]
    if driver in driver_dates_total_dic.keys():
        driver_dates_total_dic[driver].append(date)
    else:
        driver_dates_total_dic[driver] = [date]
    
    if driver in driver_dates_month_dic.keys():
        driver_dates_month_dic[driver].append(date)
    else:
        driver_dates_month_dic[driver] = [date]
        
    # month_df_total.columns = month_df_total.iloc[0]
    month_df_total = pandas.DataFrame(month_df_total.values[1:], columns=month_df_total.iloc[0])
    month_df_total = month_df_total.rename(columns={'ردیف' : 'no' , 'شماره پالت':'pallet' , 'طول':'length' , 'عرض':'width' , 'کد':'code' , 'تعداد':'count' , 'متراژ':'metters'})
    month_df_total = month_df_total[month_df_total["code"].notnull()]
    
    
    sum_of_month_billings = len(month_billings)


    for i in range(1,len(month_billings)):
        temp_df = pandas.read_excel(month_billings[i],header=1)
        driver = temp_df.columns[0]
        date = temp_df.columns[-1]
        if driver in driver_dates_total_dic.keys():
            driver_dates_total_dic[driver].append(date)
        else:
            driver_dates_total_dic[driver] = [date]
            
        if driver in driver_dates_month_dic.keys():
            driver_dates_month_dic[driver].append(date)
        else:
            driver_dates_month_dic[driver] = [date]
        
        temp_df = pandas.DataFrame(temp_df.values[1:], columns=temp_df.iloc[0])
        temp_df = temp_df.rename(columns={'ردیف' : 'no' , 'طول':'length' , 'عرض':'width' , 'کد':'code' , 'تعداد':'count' , 'متراژ':'metters'})
        temp_df = temp_df[temp_df["code"].notnull()]
        
        # month_df_total = month_df_total.append(temp_df , ignore_index=True)
        month_df_total = pandas.concat([month_df_total,temp_df], ignore_index=True)
        
    month_df_total['length'] = month_df_total['length'].astype(str)
    month_df_total['width'] = month_df_total['width'].astype(str)
    month_df_total['code'] = month_df_total['code'].astype(str)    
    month_df_total['count'] = month_df_total['count'].astype(str)
    month_df_total['metters'] = (month_df_total['metters'].astype(float).round(2)).astype(str)
      
    replaces = {'لکه':'stained' , 'اسلب':'slab' , 'درهم':'mixed' , 'مشکی':'black' , 'سانت':'4cm'}

        
        
    month_df_total.drop(columns=['no'],inplace=True)
    month_df_total["code"].replace("لکه" , "stained",inplace=True)
    month_df_total["code"].replace("لکه دار" , "stained",inplace=True)
    month_df_total["code"].replace("اسلب" , "slab",inplace=True)
    month_df_total["code"].replace("اسلب ۴" , "4cmslab",inplace=True)
    month_df_total["code"].replace("درهم" , "mixed" ,inplace=True)
    month_df_total["code"].fillna("mixed" ,inplace=True)
    month_df_total["code"].replace("مشکی" , "black" ,inplace=True)
    month_df_total["code"].replace("4سانت" , "4cm" ,inplace=True)
    month_df_total["code"].replace("4سانتی" , '4cm' ,inplace=True) 
    month_df_total["code"].replace('74' , '474' ,inplace=True)
    month_df_total["code"].replace('73' , '473' ,inplace=True)
    month_df_total["code"].replace('73a' , '473a' ,inplace=True)
    month_df_total["code"].replace('72' , '472' ,inplace=True)
    month_df_total["code"].replace('70' , '470' ,inplace=True)
    
    month_df_total["code"].replace('474A' , '474a' ,inplace=True)
    month_df_total["code"].replace('473A' , '473a' ,inplace=True)
    month_df_total["code"].replace('472A' , '472a' ,inplace=True)
    month_df_total["code"].replace('471A' , '471a' ,inplace=True)
    month_df_total["code"].replace('470A' , '470a' ,inplace=True)
    
      
    # month_df_total.to_excel(month_folder + "//total.xlsx", encoding='utf-8')
    month_metters_sum = month_df_total['metters'].astype(float).sum()
    
    
    # month_df_total_sum = month_df_total.copy()
    
    # temp = []
    # max_idx = len(month_df_total_sum)
    
    temp = []
    sum_df = pandas.DataFrame(columns = month_df_total.columns)
    max_idx = len(month_df_total)
    
    for i in range(max_idx):
        if i in temp:
            continue
        item1 = month_df_total.loc[i]
        length = item1["length"]
        width = item1['width']
        code = item1['code']
        
        temp_df = month_df_total[(month_df_total['length'] == length) & (month_df_total['width']==width) & (month_df_total['code']==code)]
        
        sum_count = temp_df['count'].astype(float).sum()
        sum_metters = temp_df['metters'].astype(float).sum()
        
        new_row = [length ,width ,code ,sum_count ,sum_metters]
        
        sum_df.loc[len(sum_df)] = new_row
        
        for index, row in temp_df.iterrows():
            temp.append(index)
       
    sum_df['width'] = sum_df['width'].astype(float)  
    sum_df['metters'] = sum_df['metters'].astype(float)  
    sum_df['width'] = sum_df['width'].astype(float)  
    
    sum_df.sort_values(by=['width' , 'metters'],ascending=False, ignore_index=True , inplace= True)
    sum_df['width'] = sum_df['width'].astype(str)
    
    sum_df.insert(5,'total_size_metters',['']*len(sum_df))
    
    temp = []
    for i in range(len(sum_df)):
              
        if i in temp:
            continue
        item = sum_df.loc[i]
        total_size_metters = item['metters']
        
        if i != len(sum_df)-1:
            j=i+1
            
            while (sum_df.loc[j]['width'] == item['width'] and sum_df.loc[j]['length'] == item['length']):    
                temp.append(j)
                total_size_metters += sum_df.loc[j]['metters']
                # print(i)
                # print(j)
                j += 1
                if j == len(sum_df):
                    break
              
        sum_df.loc[i] = [item['length'], item['width'], item['code'], item['count'], item['metters'],  total_size_metters]
        
    
    sum_df_to_save = sum_df.copy()
    sum_df_to_save.rename(columns={'length' : 'طول', 'width': 'عرض', 'code': 'کد' , 'count': 'تعداد' , 'metters':'متراژ' , 'total_size_metters': 'جمع متراژ سایز'} , inplace=True)
    # writer = pandas.ExcelWriter(month_folder + "//مجموع.xlsx")
    # sum_df_to_save.to_excel(writer, encoding='utf-8')
    # auto_adjust_xlsx_column_width(sum_df_to_save,writer ,sheet_name="Sheet1")
    
    sum_df_to_save.to_excel(month_folder + "//مجموع.xlsx")
    func_var.autofit(month_folder + "//مجموع.xlsx", True)
    month_df_total_sum = sum_df.copy()
    
    
    
    
#%% ANALYSIS OPERATION

    analysis_df_columns = ['REPORT' , 'AMOUNT']
    analysis_df = pandas.DataFrame(columns=analysis_df_columns)
    
    analysis_df.loc[len(analysis_df)] = ['متراژ کل ارسالی' , month_metters_sum]
    analysis_df.loc[len(analysis_df)] = ['تعداد بار های ارسالی' , sum_of_month_billings]
    for driver in driver_dates_month_dic.keys():
        analysis_df.loc[len(analysis_df)] = [driver , len(driver_dates_month_dic[driver])]
    
    analysis_df.loc[len(analysis_df)] = ['','']
    #calculating codes
    codes_total_count = {}

    codes_total_count['متراژ کد 470'] = month_df_total_sum[month_df_total_sum['code'] == '470']['metters'].astype(float).sum().round(2)
    codes_total_count['متراژ کد 470a'] = month_df_total_sum[month_df_total_sum['code'] == '470a']['metters'].astype(float).sum().round(2)
    
    codes_total_count['متراژ کد 471'] = month_df_total_sum[month_df_total_sum['code'] == '471']['metters'].astype(float).sum().round(2)
    codes_total_count['متراژ کد 471a'] = month_df_total_sum[month_df_total_sum['code'] == '471a']['metters'].astype(float).sum().round(2)
    
    codes_total_count['متراژ کد 472'] = month_df_total_sum[month_df_total_sum['code'] == '472']['metters'].astype(float).sum().round(2)
    codes_total_count['متراژ کد 472a'] = month_df_total_sum[month_df_total_sum['code'] == '472a']['metters'].astype(float).sum().round(2)
    
    codes_total_count['متراژ کد 473'] = month_df_total_sum[month_df_total_sum['code'] == '473']['metters'].astype(float).sum().round(2)
    codes_total_count['متراژ کد 473a'] = month_df_total_sum[month_df_total_sum['code'] == '473a']['metters'].astype(float).sum().round(2)
    
    codes_total_count['متراژ کد 474'] = month_df_total_sum[month_df_total_sum['code'] == '474']['metters'].astype(float).sum().round(2)
    codes_total_count['متراژ کد 474a'] = month_df_total_sum[month_df_total_sum['code'] == '474a']['metters'].astype(float).sum().round(2)
    
    codes_total_count['متراژ کد درهم'] = month_df_total_sum[month_df_total_sum['code'] == 'mixed']['metters'].astype(float).sum().round(2)
    codes_total_count['متراژ کد لکه دار'] = month_df_total_sum[month_df_total_sum['code'] == 'stained']['metters'].astype(float).sum().round(2)
    
    codes_total_count['متراژ اسلب'] = month_df_total_sum[month_df_total_sum['code'] == 'slab']['metters'].astype(float).sum().round(2)
    codes_total_count['متراژ اسلب 4 سانتی'] = month_df_total_sum[month_df_total_sum['code'] == '4cmslab']['metters'].astype(float).sum().round(2)
    
    codes_total_count['متراژ کد مشکی'] = month_df_total_sum[month_df_total_sum['code'] == 'black']['metters'].astype(float).sum().round(2)
    
    codes_total_count['متراژ چهارسانتی'] = month_df_total_sum[month_df_total_sum['code'] == '4cm']['metters'].astype(float).sum().round(2)
    for item in codes_total_count.keys():
        if codes_total_count[item] != 0:
            analysis_df.loc[len(analysis_df)] = [item , codes_total_count[item]]
    
    date = jdatetime.date.today()
    analysis_df.loc[len(analysis_df)] = ['' , '']
    analysis_df.loc[len(analysis_df)] = ['تاریخ اپدیت' , str(date)]
    analysis_df.to_excel(month_folder + "//گزارش ماهانه.xlsx")
    func_var.autofit(month_folder + "//گزارش ماهانه.xlsx",True)
    
    if int(month[:month.index('-')]) in monthes_to_name.keys():
        total_report_df.loc[len(total_report_df)] = ['تعداد بار ارسالی ماه '+monthes_to_name[int(month[:month.index('-')])] , sum_of_month_billings]
        total_report_df.loc[len(total_report_df)] = ['متراژ ارسالی ماه '+monthes_to_name[int(month[:month.index('-')])] , month_metters_sum]
        total_report_df.loc[len(total_report_df)] = ['','']
        
    else:        
        total_report_df.loc[len(total_report_df)] = [month , month_metters_sum]
        total_report_df.loc[len(total_report_df)] = [month , sum_of_month_billings]
    
    # total_report_df.loc(len(total_reoirt_df)) = ['متراژ ارسالی ']

for driver in driver_dates_total_dic.keys():
    total_report_df.loc[len(total_report_df)] = [driver , len(driver_dates_total_dic[driver])]

total_report_df.loc[len(total_report_df)] = ['تاریخ اپدیت' , str(date)]
total_report_df.to_excel(directory + "//گزارش کلی.xlsx")
func_var.autofit(directory + "//گزارش کلی.xlsx" , True)
