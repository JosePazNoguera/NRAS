'''
Authors: Kharesa-Kesa and Jose
This script will attempt to replicate the actions of the spreadsheet generating the all_stations tab
'''

# from asyncio.unix_events import _UnixSelectorEventLoop
# from email import header
#from macpath import split
import os
from operator import index
from matplotlib.pyplot import polar
import pandas as pd, numpy as np, glob, ast, openpyxl as xl, shutil, pyodbc, random, datetime
from openpyxl import Workbook
from openpyxl import load_workbook
from datetime import datetime

from pandas import DataFrame


def get_orr_step_free_category(search_value):
    orr = {'B': 'Bottom', 'B2': 'Bottom', 'B3': 'Middle', 'C': 'Top'}

    return orr[search_value]

def get_connectivity_journeys_matrix(search_value):
    e = {'TopTop': 1, 'TopMiddle': 2, 'TopBottom': 3, 'MiddleTop': 1, 'MiddleMiddle': 2, 'MiddleBottom': 3,
         'BottomTop': 2, 'BottomMiddle': 3, 'BottomBottom': 5}

    return e[search_value]

def get_mobility_isolation_matrix(search_value):
    e = {'TopTop': 1, 'TopMiddle': 1, 'TopBottom': 2, 'MiddleTop': 1, 'MiddleMiddle': 2, 'MiddleBottom': 3,
         'BottomTop': 2, 'BottomMiddle': 3, 'BottomBottom': 3}

    return e[search_value]

def get_list_col():
    three = ['Isolation_(1_if_no_Cat_A_in_20_mins_drive_isochrone)',
             'Additional_Flags', 'Original_Isolation_Score',
             'Revisited_Isolation_score', 'Mobility/Isolation',
             'Isolation_and_Current_Access_Matrix_Outcome', 'Socioeconomic_Flags',
             'Socioeconomic_classification', 'Local_Impact_Score',
             'Local_Impact_Classification', 'Socioecon_/_Local_Impact',
             'Socioeconomic_/_Local_Matrix_outcome', 'Average_of_two_scores',
             'Score__without_modifier', 'Base_score_+_modifiers',
             'Score_with_modifier', 'Footfall_Modifier', 'Final_Outcome', 'Change',
             'Region_and_Final_Score', 'Journeys_and_Final_Score',
             'Region_and_Local_Factor', 'Score_Change']

    return three

def get_DfT_Num_Cat(search_value):

    if search_value.isna():
        return None

    else:

        e = {'A':6, 'B':5, 'C':4, 'D':3, 'E':2, 'F':1, 'NaN': None}
        return e[search_value]

def input_OD_Matrix():
    # inputs
    # connect to the access database
    conn = pyodbc.connect(
        r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:/Users/jose.delapaznoguera/OneDrive - Arup/NRAS Secondment/Automation/Inputs/MOIRAOctober22.accdb;')

    query = 'select * from JoinCodes'
    OD_df = pd.read_sql(query, conn)

    return OD_df


def map_input_stations(OD_df, upgrade_list):
    # Add numerical score based on a dictionary
    my_dict = {'0': 0, 'A': 1, 'B': 2, 'B1': 3, 'B2': 4, 'B3': 5, 'C': 6, 'Null': -1}
    origin_score = OD_df.AfAOrigin.map(my_dict)
    destination_score = OD_df.AfADest.map(my_dict)
    OD_df['origin_score'] = origin_score
    OD_df['destination_score'] = destination_score
    OD_df['Total_Journeys'] = OD_df['STDJOURNEYS'] + OD_df['1stJOURNEYS']

    # Categorize the journeys by looking at the maximum numerical score of the origin and destination stations
    jny_score = np.maximum(OD_df.origin_score, OD_df.destination_score)

    # Add the list as a new column to the existing dataframe
    OD_df['jny_score'] = jny_score
    my_dict_2 = {1: 'A', 2: 'B', 3: 'B1', 4: 'B2', 5: 'B3', 6: 'C'}
    jny_category = OD_df.jny_score.map(my_dict_2)
    OD_df['jny_category'] = jny_category
    OD_df['concat_categories'] = OD_df.AfAOrigin + OD_df.AfADest

    # print(base_df.tail(10))

    ## TEST
    OD_df.isna().sum()

    # Identify OD pairs where the score is 0 or Null. Save the number of removed records for traceability
    dropped_origins = len(OD_df[(OD_df.origin_score < 1) | (OD_df.origin_score.isna())])
    dropped_destinations = len(OD_df[(OD_df.destination_score < 1) | (OD_df.destination_score.isna())])
    # OD_df = OD_df.drop(OD_df[(OD_df.origin_score < 1) | (OD_df.origin_score.isna())].index)
    # OD_df = OD_df.drop(OD_df[(OD_df.destination_score < 1) | (OD_df.destination_score.isna())].index)

    print(
        f"A total of {dropped_origins:n} origins and {dropped_destinations:n} destinations were invalid and not included in the analysis.")

    New_ODMatrix = OD_df.copy()

    for tlc in upgrade_list.TLC:
        if str(tlc) == 'nan':
            continue
        # Update category. It is necessary to use the .item() method to the Series ,
        new_category = upgrade_list.loc[upgrade_list.TLC == tlc, 'New_Category'].item()

        # Update origin category
        New_ODMatrix.loc[New_ODMatrix.OriginTLC == str(tlc), 'AfAOrigin'] = new_category
        # Update destination category
        New_ODMatrix.loc[New_ODMatrix.DestinationTLC == str(tlc), 'AfADest'] = new_category

    # Update origin score
    New_ODMatrix.origin_score = New_ODMatrix.AfAOrigin.map(my_dict)
    # Update destination score
    New_ODMatrix.destination_score = New_ODMatrix.AfADest.map(my_dict)
    # Update journey score
    New_ODMatrix.jny_score = np.maximum(New_ODMatrix.origin_score, New_ODMatrix.destination_score)
    # Update journey category
    New_ODMatrix.jny_category = New_ODMatrix.jny_score.map(my_dict_2)
    # Concat the 2 categories together
    New_ODMatrix['concat_categories'] = New_ODMatrix.AfAOrigin + New_ODMatrix.AfADest

    New_ODMatrix.drop(axis=1,columns=['origin_score', 'destination_score', 'jny_score'], inplace=True)

    # Create pivot table for the final output
    base_pivot = OD_df.pivot_table(index='AfAOrigin', columns='AfADest', values='Total_Journeys', aggfunc=np.sum)
    pivot = New_ODMatrix.pivot_table(index='AfAOrigin', columns='AfADest', values='Total_Journeys', aggfunc=np.sum)

    # dataframes where only the origin or the destination are accessible
    OD_df_ass_origin = OD_df
    OD_df_ass_destination = OD_df

    grouped_origin_df = (OD_df_ass_origin.groupby(["OriginTLC", "AfAOrigin"])["Total_Journeys"].sum()).to_frame()
    grouped_origin_df.reset_index(inplace=True)

    grouped_destination_df = (
        OD_df_ass_destination.groupby(["DestinationTLC", "AfADest"])["Total_Journeys"].sum()).to_frame()
    grouped_destination_df.reset_index(inplace=True)

    # setting all Cat A as None

    grouped_origin_df.loc[grouped_origin_df.AfAOrigin == 'A', 'Total_Journeys'] = None
    grouped_origin_df.loc[grouped_origin_df.AfAOrigin == 'B1', 'Total_Journeys'] = None

    grouped_destination_df.loc[grouped_destination_df.AfADest == 'A', 'Total_Journeys'] = None
    grouped_destination_df.loc[grouped_destination_df.AfADest == 'B1', 'Total_Journeys'] = None

    return grouped_origin_df, grouped_destination_df, New_ODMatrix, pivot, base_pivot


def into_stepfree_spreadsheet(grouped_origin_df, grouped_destination_df, path_of_spreadsh, scenario_desc, pivot):

    #Final_df is the all_stations sheet here with the new updated station cateogries in in
    #grouped origin and destination are grouped dfs of the total journeys grouped by station
    #this method is to write to the new spreadsheet 

    
    target = r"C:/Users/jose.delapaznoguera/OneDrive - Arup/NRAS Secondment/Automation/Step Free Scoring_JDL_v4.00_"+str(scenario)+'.xlsx'

    #copying the path to spreadsheet file as the target scenario
    shutil.copyfile(path_of_spreadsh, target)

    #reading from the spreadsheet
    st_cat_df = pd.read_excel(path_of_spreadsh, sheet_name="St_Cat", engine='openpyxl')
    st_cat_df.rename(columns={'CRS Code': 'CRS_Code', 'Station Name (MOIRA Name)': 'Station_Name'}, inplace=True)

    st_cat_df = st_cat_df.loc[:,~st_cat_df.columns.duplicated()].copy()

    kpi_df = kpi(New_ODMatrix, OD_df)

    # Next steps upgrading the data in the clone to this current scenario:
    # 1. find the stations to be upgraded.
    for code in upgrade_list.TLC:
        if str(code) == 'nan':
            continue

        #if the current code is in the station category spreadsheet then upgrade st_cat_df with new category
        if code in st_cat_df['CRS_Code'].values:
            new_category = upgrade_list.loc[upgrade_list.TLC == code, 'New_Category'].item()
            st_cat_df.loc[st_cat_df.CRS_Code == str(code), 'Including CP6 AfA'] = new_category
        else:
            continue

    #export back to Excel
    with pd.ExcelWriter(target, mode="a", engine="openpyxl", if_sheet_exists='replace') as writer:

        st_cat_df.to_excel(writer, sheet_name="St_Cat", index=False)
        scenario_desc.to_excel(writer, sheet_name="Scen_desc", index=False)
        kpi_df.to_excel(writer, sheet_name="KPI_Py", index=False)
        pivot.to_excel(writer, sheet_name="Pivot")


        grouped_origin_df.to_excel(writer, sheet_name="Inaccessible O Accessi D")
        grouped_destination_df.to_excel(writer, sheet_name="Accessible O Inaccessi D")

    #done


def output_to_log(input_df, scenario ):

    now = datetime.now()
    dt_string = now.strftime("%d/%m/%Y %H:%M:%S")

    with open(r"C:\Users\jose.delapaznoguera\OneDrive - Arup\NRAS Secondment\Automation\scenarios_output_log.txt", "a+") as file_object:
        # Move read cursor to the start of file.
        file_object.seek(0)
        # If file is not empty then append '\n'
        data = file_object.read(100)
        if len(data) > 0 :
            file_object.write("\n")
        # Append text at the end of file
        file_object.write("\n")
        line = 'Scenario number: ' + str(scenario) + ' run at ' + dt_string
        file_object.write(line)
        file_object.write("\n")
        file_object.write(str(input_df))
        file_object.write("\n")

def make_kepler_input(final_df, path_of_spreadsh, scenario):

    #naming this dataframe kepler as it will serve as the stations.csv file for kepler
    kepler = pd.read_excel(path_of_spreadsh, sheet_name="Coordinates", header=2, usecols="B:F", engine='openpyxl')

    #running over the final_df dafeframe and adding the new categories and other columns to it
    for code in final_df.Unique_Code:
        if str(code) == 'nan':
            continue
        if code in kepler.values:

            # Update category. It is necessary to use the .item() method to the Series ,
            kepler.loc[kepler.CRSCode == str(code), 'Step Free Category'] = final_df.loc[final_df.Unique_Code == code, 'ORR_Step_Free_Category'].item()
            #kepler.loc[kepler.CRSCode == str(code), 'DfT Category']  = final_df.loc[final_df.Unique_Code == code, 'Dft_Category'].item()
            kepler.loc[kepler.CRSCode == str(code), 'Journeys Unlocked']  = final_df.loc[final_df.Unique_Code == code, '2019_Total_Unlocked_Journeys'].item()
            kepler.loc[kepler.CRSCode == str(code), 'Connectivity Rank']  = final_df.loc[final_df.Unique_Code == code, '2019_Connectivity_Rank'].item()
            kepler.loc[kepler.CRSCode == str(code), 'Isolation Score']  = final_df.loc[final_df.Unique_Code == code, 'Isolation_and_Current_Access_Matrix_Outcome'].item()
            kepler.loc[kepler.CRSCode == str(code), 'Final Matrix Outcome']  = final_df.loc[final_df.Unique_Code == code, 'Final_Outcome'].item()
            #kepler.loc[kepler.CRSCode == str(code), 'DfT Category A to F']  = get_DfT_Num_Cat(final_df.loc[final_df.Unique_Code == code, 'Dft_Category'].item())

    #Exporting the file as a csv
    csv_outpath = 'Stations_sc_'+str(scenario)
    kepler.to_csv(csv_outpath)

def scenario_input():
    input_path = "C:/Users/jose.delapaznoguera/OneDrive - Arup/NRAS Secondment/Automation/Inputs/Suggested Example Scenario for Jose.xlsx"
    scen_desc = pd.read_excel(input_path, sheet_name="Scenario Master", engine='openpyxl', index_col=0)
    input_df = pd.read_excel(input_path, sheet_name="Diffs", engine='openpyxl', index_col=0)
    input_df.drop(['Station Name (MOIRA Name)', 'CP5', 'CP6'], axis=1, inplace=True)
    input_df.replace(to_replace="ignore", value=np.NaN, inplace=True)
    return scen_desc, input_df

def kpi(New_ODMatrix, OD_df):

    #KPIs for the step-free spreadsheet
    step_free_jnys = New_ODMatrix.Total_Journeys.loc[(New_ODMatrix['jny_category'] == "A") | (New_ODMatrix['jny_category'] == "B1")].sum()
    step_free_jnys_pctg = step_free_jnys / OD_df.Total_Journeys.sum()
    B3_or_C_stns = New_ODMatrix.Total_Journeys.loc[(New_ODMatrix['jny_category'] == "B3") | (New_ODMatrix['jny_category'] == "C")].sum()
    B3_or_C_stns_pctg = B3_or_C_stns / OD_df.Total_Journeys.sum()
    my_dict = {'step_free_journeys_%': [step_free_jnys_pctg], 'B3_or_C_stations_&': [B3_or_C_stns_pctg]}
    kpi_df = pd.DataFrame(data=my_dict)
    return kpi_df


#Pseudo-Main

scen_desc, input_df =scenario_input()
OD_df = input_OD_Matrix()

for scenario in input_df.columns:
    # clones spreadsheet as to not affect the original when writing to the sheet
    original = r"C:/Users/jose.delapaznoguera/OneDrive - Arup/NRAS Secondment/Automation/Step Free Scoring_JDL_v4.00.xlsx"
    clone = r"C:/Users/jose.delapaznoguera/OneDrive - Arup/NRAS Secondment/Automation/Step Free Scoring_JDL_v4.00_clone.xlsx"
    shutil.copyfile(original, clone)

    path_of_spreadsh = clone
    scenario_desc = pd.DataFrame(scen_desc.loc[scenario][['Description', 'Notes']])

    # base_df = pd.read_excel(path_of_spreadsh, sheet_name="All Stations", engine='openpyxl')
    # with the option of selecting table from sheet
    # base_df = pd.read_excel(path_of_spreadsh, sheet_name="All Stations", header=2, usecols="B:AS", engine='openpyxl')
    # alt_any = pd.read_excel(path_of_spreadsh, sheet_name="Alt_Any_20", header=4, usecols="B:F", engine='openpyxl')

    # base_df.columns = [c.replace(' ', '_') for c in base_df.columns]
    # alt_any.columns = [c.replace(' ', '_') for c in alt_any.columns]

    upgrade_list = pd.DataFrame()
    upgrade_list['TLC'] = input_df.index
    upgrade_list['New_Category'] = input_df[scenario].values
    upgrade_list.dropna(inplace=True)
    upgrade_list.reset_index()
    grouped_origin_df, grouped_destination_df, New_ODMatrix, pivot, base_pivot = map_input_stations(OD_df, upgrade_list)
    ##
    output_to_log(upgrade_list, str(scenario))
    #
    into_stepfree_spreadsheet(grouped_origin_df, grouped_destination_df, path_of_spreadsh, scenario_desc, pivot)

    print(f"Scenario {scenario} run successfully. {len(upgrade_list)} stations were upgraded")

os.remove(clone)
print(f"Process finished successfully")
