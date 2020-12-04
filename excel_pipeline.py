import numpy as np
from openpyxl import load_workbook
from os import listdir, remove
import pandas as pd

FILE_1 = 'shift-data.xlsx'
FILE_2 = 'third-shift-data.xlsx'

INPUT_PATH = 'input/'
OUTPUT_PATH = 'output/'

def merge_excel_files(filename1, filename2, output):
    """Takes in two excel files and a desired output 
        excel file name. Creates a single output excel 
        file with each sheet from the original files.

        Example: input file1 contains sheets: Products
                    and Stores
                 input file2 contains sheet: Orders
                 output file is named: Coffee Shop

                 final output will be: Coffee Shop.xlsx,
                    containing three sheets: Products,
                    Stores, and Orders 

    Args:
        filename1 (str): Name of Excel file one
        filename2 (str): Name of Excel file two
        output (str): Name of output Excel file
    """
    complete_path1 = INPUT_PATH + filename1
    complete_path2 = INPUT_PATH + filename2
    sheet_list1 = pd.ExcelFile(complete_path1).sheet_names
    sheet_list2 = pd.ExcelFile(complete_path2).sheet_names

    output_name = '{}{}.xlsx'.format(OUTPUT_PATH, output)
    writer = pd.ExcelWriter(output_name) # pylint: disable=abstract-class-instantiated

    for sheet in sheet_list1:
        data = pd.read_excel(complete_path1, sheet)
        data.to_excel(writer, sheet, index=False)
    
    for sheet in sheet_list2:
        data = pd.read_excel(complete_path2, sheet)
        data.to_excel(writer, sheet, index=False)
    
    writer.save()

def clean_output_directory():
    """Automatically checks the output directory and deletes
        any files that currently exist. This is for organization,
        only allowing files the user is expecting from the current
        run in the output folder.
    """
    if listdir(OUTPUT_PATH):
        [remove('{}{}'.format(OUTPUT_PATH,filename)) for filename in listdir(OUTPUT_PATH)]

def pivot_tables(filename, column):
    """Merges all sheets within the entered Excel file
        together, and creates a pivot table. The pivot
        table groups the rows by the Shift column and
        measures the average values per Shift. The pivot
        table is appended to the existing Excel workbook
        as a new sheet called "pivot".

    Args:
        filename (str): Excel file name
        column (str): Column name to group by
    """
    complete_path = OUTPUT_PATH + filename

    sheet_list = pd.ExcelFile(complete_path).sheet_names

    all_data = pd.DataFrame()
    for sheet in sheet_list:
        cur_data = pd.read_excel(complete_path, sheet, index=False)
        all_data = all_data.append(cur_data)
    
    
    pivot = all_data.groupby([column]).mean()
    productivity = pivot.loc[:, 'Production Run Time (Min)':'Products Produced (Units)']

    book = load_workbook(complete_path)
    writer = pd.ExcelWriter(complete_path, engine='openpyxl') # pylint: disable=abstract-class-instantiated
    writer.book = book
    productivity.to_excel(writer, 'pivot', index=False)
    writer.save()

def main():
    clean_output_directory()
    
    merge_excel_files(FILE_1, FILE_2, 'file_merge')
    print(('Successfully merged input/{} and input/{} into output/file_merge.xlsx').format(FILE_1, FILE_2, ))
    
    pivot_tables('file_merge.xlsx', 'Shift')
    print('Created pivot table in output/file_merge.xlsx, it is in a new sheet called "pivot".')
    
main()
