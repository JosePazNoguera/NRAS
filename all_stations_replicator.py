'''
Authors: Kharesa-Kesa and Jose
This script will attempt to replicate the actions of the spreadsheet generating the all_stations tab



'''


from asyncio.unix_events import _UnixSelectorEventLoop
from email import header
from operator import index
from matplotlib.pyplot import polar
import pandas as pd, numpy as np, glob, ast, openpyxl, shutil, pyodbc, random, datetime
from openpyxl import Workbook
from openpyxl import load_workbook
from datetime import datetime


def get_new_categories(base_df):

    #this method does all the calculating the new categories

    # Import stations to be upgraded
    input_path = '/Users/kharesa-kesa.spencer/Library/CloudStorage/OneDrive-Arup/Projects/Network Rail Accessibility case/matrices/Input template.csv'
    upgrade_list = pd.read_csv(input_path)


    for tlc in upgrade_list.TLC:
        if str(tlc) == 'nan':
            continue
        # Update category. It is necessary to use the .item() method to the Series ,
        new_category = upgrade_list.loc[upgrade_list.TLC == tlc, 'New_Category'].item()

        if tlc in base_df.values:
            # Update ORR_Step_Free_Category
            base_df.loc[base_df.Unique_Code == str(tlc), 'ORR_Step_Free_Category'] = new_category

    base_df['Inaccessible_(1_if_not_Step_Free_Cat._A)'] = 1
    base_df.loc[(base_df.ORR_Step_Free_Category == 'A') | (base_df.ORR_Step_Free_Category == 'B1'), 'Inaccessible_(1_if_not_Step_Free_Cat._A)'] = 0
    base_df.loc[base_df.ORR_Step_Free_Category.isna(), 'Inaccessible_(1_if_not_Step_Free_Cat._A)'] = np.NaN

    
    return base_df


def set_connectivity(updated_cats):

    updated_cats




def main():
    #any variables for the main 

    path_of_spreadsh = '/Users/kharesa-kesa.spencer/Library/CloudStorage/OneDrive-Arup/Projects/Network Rail Accessibility case/CSV WORK/Step Free Scoring_JDL_v3.00_clone.xlsx'
    base_df  = pd.read_excel(path_of_spreadsh, sheet_name = "All Stations", header=2 , usecols="B:AS", engine='openpyxl')
    base_df.columns = [c.replace(' ','_') for c in base_df.columns]


    updated_cats = get_new_categories(base_df)
    updated_jnys = set_connectivity(updated_cats)




if __name__ == "__main__":
    main()
