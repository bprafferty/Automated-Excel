Title: Automated Excel Data Pipeline

Author: Brian Rafferty

Date: 12/4/2020

Description: This Python script automates the process
    of merging multiple Excel files into the same
    Workbook. Each individual sheet's data and title
    is preserved in the output file. 
    
    After merging files, the script takes the new excel 
    workbook and creates a single pivot table that analyzes
    every sheet. In this case, the pivot table groups
    the data by Shift and then takes the average of
    the rest of the columns for each shift. By doing
    so, average productivity per shift can be calculated.

    The goal is to save a Data Analyst a substantial amount
    of time when working with Excel files.

Wish List: This script can be augmented by adding user 
    interface functionality. Doing so would allow users to 
    customize which Excel files are merged, which column(s) 
    are selected to group-by within pivot tables, and which 
    aggregate functions to apply.


Dependencies:
    - Numpy
    - Openpyxl
    - OS
    - Pandas
    - Python 3.8
