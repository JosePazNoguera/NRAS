'''
Authors: Kharesa-Kesa and Jose
This script will attempt to replicate the actions of the spreadsheet generating the all_stations tab



'''


from email import header
from operator import index
from matplotlib.pyplot import polar
import pandas as pd, numpy as np, glob, ast, openpyxl, shutil, pyodbc, random, datetime
from openpyxl import Workbook
from openpyxl import load_workbook
from datetime import datetime







def main():
    #any variables for the main 

    path_of_spreadsh = '/Users/kharesa-kesa.spencer/Library/CloudStorage/OneDrive-Arup/Projects/Network Rail Accessibility case/CSV WORK/Step Free Scoring_JDL_v3.00_clone.xlsx'
    base_df  = pd.read_excel(path_of_spreadsh, sheet_name = "All Stations", header=2 , usecols="B:AS", engine='openpyxl')
    




if __name__ == "__main__":
    main()
