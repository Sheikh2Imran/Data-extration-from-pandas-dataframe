import pandas as pd
from functools import reduce
from pandas import read_excel
from pandas import ExcelWriter

from fuzzywuzzy import fuzz
from fuzzywuzzy import process

class Employee:
    def read_xl_file(self, file_name):
        return read_excel(file_name)
    
    def preprocess_data_with_regex(self, dataframe):
        regex1 = '\(.*?\)'
        regex2 = '\d'         
        regex3 = '\w-'
        regex4 = '^md.'
        dataframe['Name'].replace(regex=True, inplace=True, to_replace=[regex1, regex2, regex3, regex4], value=r'')
        return dataframe
    
    def make_name_lower_case(self, dataframe):
        dataframe['name'] = dataframe['Name']
        dataframe['name'] = dataframe['name'].str.lower()
        return dataframe

    def remove_white_space(self, dataframe):
        dataframe['name'] = dataframe['name'].str.replace(' ', '')
        return dataframe
    
    def fuzzy_merge_df1_and_df2_name(self, df_1, df_2):
        df_1.loc[:, 'Device ID'] = ''
        df_1.loc[:, 'Name_Device'] = ''
        df_2.loc[:, 'Name'] = ['Null' if name2 is '' else name2 for name2 in df_2['Name']]
        
        for i in range(len(df_1)):
            for j in range(len(df_2)):
                if df_1.loc[i, 'Name'].lower().replace(' ', '') == df_2.loc[j, 'Name'].lower().replace(' ', ''):
                    df_1.loc[i, 'Device ID'] = df_2.loc[j, 'Identification No']
                    df_1.loc[i, 'Name_Device'] = df_2.loc[j, 'Name']
                    
        df1_and_df2.loc[:, 'Device ID'] = ['Null' if device_id is '' else device_id for device_id in df1_and_df2['Device ID']]
        return df_1
    
    def fuzzy_merge_df1_df2_df3_name(self, df1_and_df2, df_3):
        df1_and_df2.loc[:, 'Employee ID'] = ''
        df1_and_df2.loc[:, 'Name_Employee'] = ''
        
        for i in range(len(df1_and_df2)):
            for j in range(len(df_3)):
                if df1_and_df2.loc[i, 'Name'].lower().replace(' ', '') == df_3.loc[j, 'Name'].lower().replace(' ', ''):
                    df1_and_df2.loc[i, 'Employee ID'] = df_3.loc[j, 'Id']
                    df1_and_df2.loc[i, 'Name_Employee'] = df_3.loc[j, 'Name']
#                    
        df1_and_df2.loc[:, 'Employee ID'] = ['Null' if employee_id is '' else employee_id for employee_id in df1_and_df2['Employee ID']]
        return df1_and_df2
    
    def write_dataframe_to_exel(self, dataframe):
        writer = ExcelWriter('EmployeeExport.xlsx')
        dataframe.to_excel(writer,'Sheet')
        writer.save()   



if __name__ == '__main__':
    
    employee = Employee()

    file_name1 = 'Raw Employee(odoo).xlsx'
    df1 = employee.read_xl_file(file_name1)
    df1 = employee.make_name_lower_case(df1)
    df1 = employee.preprocess_data_with_regex(df1)
    df1 = employee.remove_white_space(df1)
    
    file_name2 = 'Device Employee.xlsx'
    df2 = employee.read_xl_file(file_name2)
    df2 = employee.preprocess_data_with_regex(df2)
    df2 = employee.make_name_lower_case(df2)
    df2 = employee.remove_white_space(df2)
    
    file_name3 = 'all employee list.xls'
    df3 = employee.read_xl_file(file_name3)
    df3 = employee.preprocess_data_with_regex(df3)
    df3 = employee.make_name_lower_case(df3)
    df3 = employee.remove_white_space(df3)
    
    df1_and_df2 = employee.fuzzy_merge_df1_and_df2_name(df1, df2)
    final_dataframe = employee.fuzzy_merge_df1_df2_df3_name(df1_and_df2, df3)
    
    employee.write_dataframe_to_exel(final_dataframe)
