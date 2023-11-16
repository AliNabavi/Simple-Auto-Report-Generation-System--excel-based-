# -*- coding: utf-8 -*-
"""
Created on Sun Oct 22 11:10:41 2023

@author: Negin Eram
"""

import os,xlwings, pandas, jdatetime, pathlib


#%% SETTINGS
ignore_items = ['oil']

PROGRAM_LANGUAGE = 'eng'  #eng/per



#%% General VARS
directory = os.getcwd()
date = jdatetime.date.today()

total_report_df_columns = ['REPORT' , 'AMOUNT']

monthes_to_name = {}
if PROGRAM_LANGUAGE == 'per':
    monthes_to_name[1] = 'فروردین'
    monthes_to_name[2] = 'اردیبهشت'
    monthes_to_name[3] = 'خرداد'
    monthes_to_name[4] = 'تیر'
    monthes_to_name[5] = 'مرداد'
    monthes_to_name[6] = 'شهریور'
    monthes_to_name[7] = 'مهر'
    monthes_to_name[8] = 'ابان'
    monthes_to_name[9] = 'اذر'
    monthes_to_name[10] = 'دی'
    monthes_to_name[11] = 'بهمن'
    monthes_to_name[12]= 'اسفند'
    
else:
    monthes_to_name[1] = 'JANUARY'
    monthes_to_name[2] = 'FEBRUARY'
    monthes_to_name[3] = 'MARCH'
    monthes_to_name[4] =  'APRIL'
    monthes_to_name[5] = 'MAY'
    monthes_to_name[6] = 'JUNE'
    monthes_to_name[7] = 'JULY'
    monthes_to_name[8] = 'AUGUST'
    monthes_to_name[9] = 'SEPTEMBER'
    monthes_to_name[10] = 'OCTOBER'
    monthes_to_name[11] = 'NOVEMBER'
    monthes_to_name[12] = 'DECEMBER'




#%%File reading VARS
bill_driver_name_column = 0
bill_car_column = 1
bill_service_number_column = 3
bill_date_column = 5





#%% CLASSES

# class Driver():
#     instances = []
#     def __init__(self, driver_name, cargoes_count=0):
#         self.name = driver_name
#         self.cargoes_count = cargoes_count
#         self.instances.append(self)
        
#     def __str___(self):
#         return self.name
    
    
class Bill_entry():
    def __init__(self, bill_entry_dict):
        for i in ['length', 'width', 'code', 'count', 'meterage']:
            assert i in list(bill_entry_dict.keys()) 
    
        self.length= bill_entry_dict.get('length')
        self.width= bill_entry_dict.get('width')
        self.code= bill_entry_dict.get('code')
        self.count= bill_entry_dict.get('count')
        self.meterage= bill_entry_dict.get('meterage')
        
          

class Bill():
    instances = []
    def __init__(self, date, car, driver_name, service_number, bill_entries):
        self.date = date
        self.service_number = service_number
        self.car = car
        
        self.driver = driver_name
        
        assert type(bill_entries) == list
        self.bill_entries = [Bill_entry(bill_entry) for bill_entry in bill_entries]
        
        self.instances.append(self)



#%% Gneral FUNCTIONS 


def GET_MONTHES_LIST():
    monthes_list =[]
    temp = next(os.walk('.'))[1]
    for item in temp:
        if item[0].isdigit():
            monthes_list.append(pathlib.Path(directory+'/' + item))
    monthes_list.sort()
    
    return monthes_list



def autofit(file_path,pdf_save=False):  
    # with xlwings.App(visible=False) :
    book = xlwings.Book(file_path)
    for sheet in book.sheets:
    # sheet = book.sheets[0]
        sheet.autofit()
    book.save(file_path)
    
    if pdf_save:
        sheet = book.sheets[-1]
        sheet.api.PageSetup.Orientation = xlwings.constants.PageOrientation.xlLandscape
        book.to_pdf(file_path[:file_path.index('.')]+'.pdf', include=book.sheet_names[-1])
    
    book.close()


def SAVE_DFDICT_TO_EXCELL_SHEETS(dfs_dict, xlsx_path,pdf_save=False):
    # with pandas.ExcelWriter(xlsx_path,engine="openpyxl") as writer:
    writer = pandas.ExcelWriter(xlsx_path) 
    a = list(dfs_dict.keys())
    a.sort()
    for sheet_name in a:
        df = dfs_dict[sheet_name]
        
        df = df.style.set_properties(**{
        'font-size': '11pt',
        'text-align': 'center'})
        
        df.to_excel(writer, sheet_name=sheet_name, index=False, na_rep='')
        data = df.data
        if not data.empty:
            for column in data:
                column_length = max(data[column].astype(str).map(len).max(), len(column))
                col_idx = data.columns.get_loc(column)
                writer.sheets[sheet_name].set_column(col_idx, col_idx, column_length)
        
    writer.close()    
            
    if pdf_save:
        df = dfs_dict[list(dfs_dict.keys())[-1]]
        
        df = df.style.set_properties(**{
        'font-size': '16pt',
        'border-bottom': '1pt solid gray'
        })
        # df = df.style.background_gradient(axis=None,vmin=1, vmax=5, cmap="YlGnBu")
        
        html_path = xlsx_path[:xlsx_path.index('.')]+'.html'
        pdf_path = xlsx_path[:xlsx_path.index('.')]+'.pdf'
        
        f = open(html_path ,'w')
        a = df.to_html()
        with open(html_path, "w", encoding="utf-8") as file:
            file.writelines('<meta charset="UTF-8">\n')
            file.write(a)
        
        # pdfkit.from_file(html_path, pdf_path)


# def SAVE_DFDICT_TO_EXCELL_SHEETS(dfs_dict, xlsx_path,pdf_save=False):
#     #dfs dict with keys of sheet name and value the dataframe
#     with pandas.ExcelWriter(xlsx_path,engine="openpyxl") as writer:
#         for sheet_name in dfs_dict.keys():
#             dfs_dict[sheet_name].to_excel(writer, sheet_name= sheet_name)
#             # auto_adjust_xlsx_column_width(dfs_dict[sheet_name], writer, sheet_name=sheet_name, margin=3)
            
#     autofit(xlsx_path,pdf_save)            




# def SAVE_DFDICT_TO_EXCELL_SHEETS(dfs_dict, xlsx_path,pdf_save=False):
#     with pandas.ExcelWriter(xlsx_path,engine="openpyxl") as writer:
#         for sheet_name in dfs_dict.keys():
#             dfs_dict[sheet_name].to_excel(writer, sheet_name= sheet_name)
#             # auto_adjust_xlsx_column_width(dfs_dict[sheet_name], writer, sheet_name=sheet_name, margin=3)
            
    # autofit(xlsx_path,pdf_save)            




def ASSIGN_SIZE_COLUMN(trade_df):
    size_column= []
    
    for i in range(len(trade_df)):
        length = str(trade_df.loc[i]['length'])
        width = str(trade_df.loc[i]['width'])
        if length[-2:] == '.0':
            length = str(int(trade_df.loc[i]['length']))
        if width[-2:] == '.0':
            width = str(int(trade_df.loc[i]['width']))
        
        if width.isdigit():
            size_column.append(length + ' * ' + width)
            
        else:
            size_column.append('')
    
    trade_df = trade_df.assign(size=size_column)
    
    return trade_df



#%% Special case FUNCTIONS

def GET_BILL_FILES(folder_name):
    bill_files = []
    for file in os.listdir(folder_name):    
        if file[:3].isdigit():
            bill_files.append(pathlib.Path(str(folder_name)+'/' + file))
    return bill_files



def READ_BILL_FILE(bill_file_path):
    bill_df = pandas.read_excel(bill_file_path,header=1)
    driver_name = bill_df.columns[bill_driver_name_column]
    car = bill_df.columns[bill_car_column]
    service_number = bill_df.columns[bill_service_number_column]
    date = bill_df.columns[bill_date_column]
    
    bill_df = pandas.read_excel(bill_file_path,header=2)
    if PROGRAM_LANGUAGE == 'per':
        bill_df = bill_df[bill_df['طول'].notnull()]
    else:
        bill_df = bill_df[bill_df['width'].notnull()]
        
    #create bill class
    bill_entries = []
    for idx , entry in bill_df.iterrows():
        if type(entry.get('طول')) == str or type(entry.get('width')) ==str:
            continue
        if PROGRAM_LANGUAGE == 'per':
            bill_entry_dict = {
                'length': entry.get('طول'),
                'width':entry.get('عرض'),
                'code':entry.get('کد'),
                'count':entry.get('تعداد'),
                'meterage':entry.get('متراژ')}
            bill_entries.append(bill_entry_dict)
        
        else:
            bill_entry_dict = {
                'length': entry.get('length'),
                'width':entry.get('width'),
                'code':entry.get('code'),
                'count':entry.get('count'),
                'meterage':entry.get('meterage')}
            bill_entries.append(bill_entry_dict)
        
    bill = Bill(date, car, driver_name, service_number, bill_entries)
    return bill
    


    
def SUM_BILLS(bills_list):
    sum_df_column = ['length', 'width', 'code', 'count', 'meterage', 'total_size_meterage']
    sum_df = pandas.DataFrame(columns=sum_df_column)
    
    for bill in bills_list:
        for bill_entry in bill.bill_entries:
            temp = [bill_entry.length, bill_entry.width, bill_entry.code, bill_entry.count, bill_entry.meterage,'']
            sum_df.loc[len(sum_df)] = temp
            
    sum_df.sort_values(['width','length','code'],ascending=False,inplace=True, ignore_index=True)
    
    sum_df = CLEAR_SUM_DF_COLUMNS(sum_df)
    sum_df['size'] = sum_df['length'].astype(int).astype(str) +' * ' +sum_df['width'].astype(int).astype(str)
    sum_df = sum_df.iloc[:,[0,1,6,2,3,4,5]]
    
    sum_df = ASSIGN_SIZE_METERAGE_COLUMN(sum_df)
    return sum_df




def CLEAR_SUM_DF_COLUMNS(entry_df):
    entry_df["code"].replace("لکه" , "لکه دار",inplace=True)
    entry_df["code"].replace("اسلب" , "slab",inplace=True)
    entry_df["code"].replace("اسلب ۴" , "4cmslab",inplace=True)
    entry_df["code"].fillna("درهم" ,inplace=True)
    entry_df["code"].replace('74' , '474' ,inplace=True)
    entry_df["code"].replace('73' , '473' ,inplace=True)
    entry_df["code"].replace('73a' , '473a' ,inplace=True)
    entry_df["code"].replace('72' , '472' ,inplace=True)
    entry_df["code"].replace('70' , '470' ,inplace=True)
    
    entry_df["code"].replace('474A' , '474a' ,inplace=True)
    entry_df["code"].replace('473A' , '473a' ,inplace=True)
    entry_df["code"].replace('472A' , '472a' ,inplace=True)
    entry_df["code"].replace('471A' , '471a' ,inplace=True)
    entry_df["code"].replace('470A' , '470a' ,inplace=True)
    
    return entry_df


def ASSIGN_SIZE_METERAGE_COLUMN(entry_df):
    summery_df_columns = ['size','code','count','meterage','total_size_meterage']
    summery_df = pandas.DataFrame(columns=summery_df_columns)
    
    for size in entry_df['size'].unique():
        size_df = entry_df[entry_df['size'] == size].reset_index()
        size_meterage = size_df['meterage'].sum()
        
        t=0
        for code in size_df['code'].unique():
            size_code_df = size_df[size_df['code'] == code]
            size_code_meterage = size_code_df['meterage'].sum()
            size_code_count = size_code_df['count'].sum()
            
            if t==0:
                temp= [size,code,size_code_count,size_code_meterage, size_meterage]
            else:
                temp= [size,code,size_code_count,size_code_meterage, '']
        
            summery_df.loc[len(summery_df)] = temp
            t+=1
    return summery_df



def CALCULATE_REPORT(bills_list, sum_df, total_report_df=None, month_name=None):
    driver_bills_dict={}
    for bill in bills_list:
        if bill.driver in driver_bills_dict.keys():
            driver_bills_dict[bill.driver].append(bill)
        else:
            driver_bills_dict[bill.driver] = [bill]
    
    report_df_cols= ['Report', 'Amount']
    report_df = pandas.DataFrame(columns=report_df_cols)
   
    if PROGRAM_LANGUAGE == 'per':
        total_meterage = sum_df['meterage'].sum()
        temp = ['متراژ ارسالی ماه ' + month_name, total_meterage]
        report_df.loc[len(report_df)] = temp
            
        daily_meterage = sum_df['meterage'].sum()//date.day
        temp = [' میانگین ارسال روزانه', daily_meterage]
        report_df.loc[len(report_df)] = temp
        
        bills_count = len(bills_list)
        temp = [' تعداد بار ارسالی', bills_count]
        report_df.loc[len(report_df)] = temp
        
        report_df.loc[len(report_df)] = ['','']
        
        report_df.loc[len(report_df)] = ['','']
        report_df.loc[len(report_df)] = ['تاریخ اپدیت' , date]
        
    else:
        total_meterage = sum_df['meterage'].sum()
        temp = ['total_meterage_of_month_' + month_name, total_meterage]
        report_df.loc[len(report_df)] = temp
            
        daily_meterage = sum_df['meterage'].sum()//date.day
        temp = ['average_daily_meterage', daily_meterage]
        report_df.loc[len(report_df)] = temp
        
        bills_count = len(bills_list)
        temp = ['sum_of_month_bills', bills_count]
        report_df.loc[len(report_df)] = temp
        
        report_df.loc[len(report_df)] = ['','']
        
        report_df.loc[len(report_df)] = ['','']
        report_df.loc[len(report_df)] = ['update_date' , date]
        
    for driver in driver_bills_dict:
        temp=[driver,len(driver_bills_dict[driver])]
        report_df.loc[len(report_df)] = temp
        
        # if driver in total_report_df['Report']:
        #     total_report_df[total_report_df['Report']==driver] += 1
        # else:
            
    
    
    if month_name != None and type(total_report_df) is pandas.DataFrame:
        if PROGRAM_LANGUAGE =='per':
            total_report_df.loc[len(total_report_df)] = ['متراژ ارسالی ماه' +' ' + month_name , total_meterage]
            total_report_df.loc[len(total_report_df)] = [' میانگین ارسال روزانه', daily_meterage]
     
            total_report_df.loc[len(total_report_df)] = ['تعداد بار ارسالی ماه'+ ' ' + month_name, bills_count]
            total_report_df.loc[len(total_report_df)] = ['','']
         
        else:
             total_report_df.loc[len(total_report_df)] = ['total_production_meterage_month_' + month_name , total_meterage]
             total_report_df.loc[len(total_report_df)] = ['average_daily_meterage', daily_meterage]
      
             total_report_df.loc[len(total_report_df)] = ['number_of_billings_month' + month_name, bills_count]
             total_report_df.loc[len(total_report_df)] = ['','']
             
        
    
    return report_df, total_report_df



def clear_files(monthes_list):
    for month_folder in monthes_list:
        month_bill_files = GET_BILL_FILES(month_folder)
        
    
    
    
    
    
    
    
    
    
    
