'''
Authors: Kharesa-Kesa and Jose
This script will attempt to replicate the actions of the spreadsheet generating the all_stations tab



'''


from operator import index
from matplotlib.pyplot import polar
import pandas as pd, numpy as np, glob, ast, openpyxl, shutil, pyodbc, random, datetime
from openpyxl import Workbook
from openpyxl import load_workbook
from datetime import datetime


def resizing_spreadsh(base_df):
    
    #dropping extras
    #*TO CHANGE* if the sheet is resized and the table starts in the correct
    base_df.drop(index=0, inplace=True)
    base_df.drop(columns='Unnamed: 0', inplace=True)
    base_df.columns = base_df.iloc[0]
    base_df.drop(index=1, inplace=True)
    base_df.reset_index(inplace=True)
    base_df.drop(columns='index', inplace=True)

    base_df = base_df.loc[:,~base_df.columns.duplicated()].copy()


    return base_df






def main():
    #any variables for the main 

    path_of_spreadsh = '/Users/kharesa-kesa.spencer/Library/CloudStorage/OneDrive-Arup/Projects/Network Rail Accessibility case/CSV WORK/Step Free Scoring_JDL_v3.00_clone.xlsx'
    base_df  = pd.read_excel(path_of_spreadsh, sheet_name = "All Stations", engine='openpyxl')
    
    base_df = resizing_spreadsh(base_df)


if __name__ == "__main__":
    main()
