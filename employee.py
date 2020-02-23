import pandas as pd
from pandas import read_excel
from functools import reduce
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
        dataframe['Name'].replace(regex=True, inplace=True, to_replace=[regex1, regex2, regex3], value=r'')
        return dataframe
    
    def make_name_lower_case(self, dataframe):
        dataframe['Name'] = dataframe['Name'].str.lower()
        return dataframe

    def remove_white_space(self, dataframe):
        dataframe['Name'] = dataframe['Name'].str.replace(' ', '')
        return dataframe
    
    def fuzzy_merge_df1_and_df2(self, df_1, df_2, key1, key2, threshold=50, limit=2):
        """
        df_1 is the left table to join
        df_2 is the right table to join
        key1 is the key column of the left table
        key2 is the key column of the right table
        threshold is how close the matches should be to return a match, based on Levenshtein distance
        limit is the amount of matches that will get returned, these are sorted high to low
        """
        s = df_2[key2].tolist()
    
        m = df_1[key1].apply(lambda x: process.extract(x, s, limit=limit))    
        df_1['matches'] = m
    
        m2 = df_1['matches'].apply(lambda x: ', '.join([i[0] for i in x if i[1] >= threshold]))
        df_1['Device ID'] = df_2['Identification No']
        df_1['matches'] = m2
    
        return df_1
    
    def fuzzy_merge_df1_df2_df3(self, df_1, df_2, key1, key2, threshold=50, limit=2):
        """
        df_1 is the left table to join
        df_2 is the right table to join
        key1 is the key column of the left table
        key2 is the key column of the right table
        threshold is how close the matches should be to return a match, based on Levenshtein distance
        limit is the amount of matches that will get returned, these are sorted high to low
        """
        s = df_2[key2].tolist()
    
        m = df_1[key1].apply(lambda x: process.extract(x, s, limit=limit))    
        df_1['matches'] = m
    
        m2 = df_1['matches'].apply(lambda x: ', '.join([i[0] for i in x if i[1] >= threshold]))
        df_1['Employee ID'] = df_2['Id']
        df_1['matches'] = m2
    
        return df_1
    
    def write_dataframe_to_exel(self, dataframe):
        writer = ExcelWriter('EmployeeExport.xlsx')
        dataframe.to_excel(writer,'Sheet')
        writer.save()        
       

if __name__ == '__main__':
    
    employee = Employee()

    file_name1 = 'Raw Employee(odoo).xlsx'
    df1 = employee.read_xl_file(file_name1)
    df1 = employee.preprocess_data_with_regex(df1)
    df1 = employee.make_name_lower_case(df1)
    df1 = employee.remove_white_space(df1)
    
    file_name2 = 'Device Employee.xlsx'
    df2 = employee.read_xl_file(file_name2)
    df2 = employee.preprocess_data_with_regex(df2)
    df2 = employee.make_name_lower_case(df2)
    df2 = employee.remove_white_space(df2)
    
    # match dataframes by name
    df1_and_df2 = employee.fuzzy_merge_df1_and_df2(df1, df2, 'Name', 'Name', threshold=50)
        
    file_name3 = 'all employee list.xls'
    df3 = employee.read_xl_file(file_name3)
    df3 = employee.preprocess_data_with_regex(df3)
    df3 = employee.make_name_lower_case(df3)
    df3 = employee.remove_white_space(df3)
    
    # match dataframes by name
    employee_dataframe = employee.fuzzy_merge_df1_df2_df3(df1_and_df2, df3, 'Name', 'Name', threshold=50)
    
    # write pandas dataframe to exel file.
    employee.write_dataframe_to_exel(employee_dataframe)
    
    
    
    
    
    
    
