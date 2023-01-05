from asyncio import subprocess
import pandas as pd
import time
import win32com.client
import subprocess


file = "<Your file path>"

def ultima_linha(file,i):
    df = pd.read_excel(file) # Reading excel file (I was using a xlsm file)

    last_row = df.loc[int(len(df['Title']))-1] # Searching the value by row ID. Using len() to get the column max lenght

    last_row_title = last_row['Title'] # Getting just one value of my last row


    #Using this code just to identify an information before and after update
    if i == 1: 
        print('Último ID antes da atualização: '+str(last_row_title))
    else:
        print('Último ID depois da atualização: '+str(last_row_title))

# I had to use the ultima_linhas' code twice, so I created the function
ultima_linha(file=file,i=1)

# Executing the process to update the table in excel file
excel = win32com.client.DispatchEx("Excel.Application") # Opening excel application
workbook = excel.Workbooks.Open(file) # Opening specific file
try:
    workbook.RefreshAll() # Function to refresh all tables
    excel.CalculateUntilAsyncQueriesDone() # IDK what is this
    workbook.Save() # Saving file



    # Issue with this code. It isn't allowed to quit or close excel file
    excel.Quit() # Closing file




except PermissionError:
    subprocess.call('TASKKILL /F /IM excel.exe')

ultima_linha(file=file,i=2) # Executing again to check the last value
