# -*- coding: utf-8 -*-
"""
Created on Sun Nov  5 10:50:49 2023

@author: Negin Eram
"""

import time,path,pandas,numpy,matplotlib,os,re,pathlib,jdatetime, xlrd ,jpype, asposecells, jdatetime, pathlib
import func_var 

# jpype.startJVM()
# from asposecells.api import Workbook


#%%

# time.sleep(60)


directory = func_var.directory


monthes_list = func_var.GET_MONTHES_LIST()
total_bills = []

total_report_df_cols= ['Report', 'Amount']
total_report_df = pandas.DataFrame(columns=total_report_df_cols)

total_report_dic = {}
t=1
for month_folder in monthes_list:
    month_name = month_folder.name
    month_name = month_name[3:]
    
    month_bills = []
    month_bill_files = func_var.GET_BILL_FILES(month_folder)
    for bill_file in month_bill_files:
        bill = func_var.READ_BILL_FILE(bill_file)
        
        month_bills.append(bill)
        total_bills.append(bill)
        
    sum_df = func_var.SUM_BILLS(month_bills)
    report_df,total_report_df = func_var.CALCULATE_REPORT(month_bills, sum_df, total_report_df,month_name)
    
    if func_var.PROGRAM_LANGUAGE == 'per': 
        func_var.SAVE_DFDICT_TO_EXCELL_SHEETS({'مجموع ارسالات' : sum_df,'گزارش':report_df}, pathlib.Path(str(month_folder) + '/گزارش-ماه.xlsx'))       
        total_report_dic[str(t) + '1 - ' + 'گزارش کلی ماه ' + month_name] = report_df
        total_report_dic[str(t) + '2 - ' + 'مجموع ارسال ماه ' + month_name] = sum_df
        
    else:
        func_var.SAVE_DFDICT_TO_EXCELL_SHEETS({'total production' : sum_df,'sum-report':report_df}, pathlib.Path(str(month_folder) + '/month_report.xlsx')) 
        total_report_dic[str(t) + '1_' + 'total_of_month_' + month_name] = report_df
        total_report_dic[str(t) + '2_' + 'sum_of_month_' + month_name] = sum_df
    
    
    t+=1

total_df = func_var.SUM_BILLS(total_bills)


driver_bills_dict={}
for bill in total_bills:
    if bill.driver in driver_bills_dict.keys():
        driver_bills_dict[bill.driver].append(bill)
    else:
        driver_bills_dict[bill.driver] = [bill]
for driver in driver_bills_dict:
    temp=[driver,len(driver_bills_dict[driver])]
    total_report_df.loc[len(total_report_df)] = temp

if func_var.PROGRAM_LANGUAGE == 'per': 
    total_report_df.loc[len(total_report_df)] = ['تاریخ اپدیت', func_var.date]
    total_report_dic['02-مجموع ارسال کل'] = total_df
    total_report_dic['01-گزارش خلاصه'] = total_report_df
    func_var.SAVE_DFDICT_TO_EXCELL_SHEETS(total_report_dic, pathlib.Path(directory + '/گزارش عملکرد ماهانه.xlsx'))        

    
else:
    total_report_df.loc[len(total_report_df)] = ['update_date', func_var.date]
    total_report_dic['02-total_production'] = total_df
    total_report_dic['01-sum_report_total'] = total_report_df
    func_var.SAVE_DFDICT_TO_EXCELL_SHEETS(total_report_dic, pathlib.Path(directory + '/total_report_generated.xlsx'))        
    


    
    
    
    
