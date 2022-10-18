'''
Authors: Kharesa-Kesa and Jose
This script will attempt to replicate the actions of the spreadsheet generating the all_stations tab
'''


import os

import pandas as pd, numpy as np, glob, ast, openpyxl as xl, shutil, pyodbc, random, datetime
import win32com.client as win32

from datetime import datetime

## GLOBAL VARIABLES ###

# Numerical score based on a dictionary
my_dict = {'0': 0, 'A': 1, 'B': 2, 'B1': 3, 'B2': 4, 'B3': 5, 'C': 6, 'Null': -1}

# second mapping to categorise the journeys. The journey category represents the worst score of the origin
# and destination scores
my_dict_2 = {1: 'A', 2: 'B', 3: 'B1', 4: 'B2', 5: 'B3', 6: 'C'}

# Access DB
access_DB = r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:/Users/jose.delapaznoguera/OneDrive - Arup/NRAS Secondment/Automation/Inputs/MOIRAOctober22.accdb;'

# input_df
input_path = "C:/Users/jose.delapaznoguera/OneDrive - Arup/NRAS Secondment/Automation/Inputs/Suggested Example Scenario for Jose.xlsx"

original = r"C:/Users/jose.delapaznoguera/OneDrive - Arup/NRAS Secondment/Automation/Step Free Scoring_JDL_v4.10.xlsx"
clone = r"C:/Users/jose.delapaznoguera/OneDrive - Arup/NRAS Secondment/Automation/Step Free Scoring_JDL_v4.10_clone.xlsx"

# Output to log
output_log_path = r"C:\Users\jose.delapaznoguera\OneDrive - Arup\NRAS Secondment\Automation\scenarios_output_log.txt"

# Output csv folder path (for Kepler)
kepler_path = r"C:\Users\jose.delapaznoguera\OneDrive - Arup\NRAS Secondment\Automation\Kepler_layers"

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

def input_OD_Matrix(st_cat_df):
    """ This method creates the Origin-Destination matrix using the Access database as an input.
    The method also merges the regions for each origin station using the Three Letter Code from the st_cat_df, a
    dataframe created from the St_Cat tab in the Step-free scoring spreadsheet.
    ...

    :param st_cat_df: Table containing all station categories in the base scenario (Control Period 6).
    """

    # connect to the access database
    conn = pyodbc.connect(access_DB)

    query = 'select * from JoinCodes'
    OD_df = pd.read_sql(query, conn)

    # join region column from the st_cat_df
    region = st_cat_df[['CRS_Code','Region']]

    OD_df = pd.merge(left=OD_df, right=region, how='outer', left_on='OriginTLC', right_on='CRS_Code')
    missing_regions = len(OD_df.OriginTLC.loc[OD_df.Region.isna()].unique())
    print(f'{missing_regions} origins could not be mapped to any region')

    # Add Total journeys: sum of standard and 1st class
    OD_df['Total_Journeys'] = OD_df['STDJOURNEYS'] + OD_df['1stJOURNEYS']

    origin_score = OD_df.AfAOrigin.map(my_dict)
    destination_score = OD_df.AfADest.map(my_dict)
    OD_df['origin_score'] = origin_score
    OD_df['destination_score'] = destination_score

    # Categorize the journeys by looking at the maximum numerical score (worst) of the origin and destination stations
    jny_score = np.maximum(OD_df.origin_score, OD_df.destination_score)
    OD_df['jny_score'] = jny_score

    jny_category = OD_df.jny_score.map(my_dict_2)
    OD_df['jny_category'] = jny_category
    OD_df['concat_categories'] = OD_df.AfAOrigin + OD_df.AfADest

    ## TEST --> Look at tests.py
    OD_df.isna().sum()

    # Identify OD pairs where the score is 0 or Null. Save the number of removed records for traceability
    dropped_origins = len(OD_df.OriginTLC.loc[(OD_df.origin_score < 1) | (OD_df.origin_score.isna())].unique())
    dropped_destinations = len(OD_df.DestinationTLC.loc[(OD_df.destination_score < 1) | (OD_df.destination_score.isna())].unique())

    # Removal of OD pairs where score is 0 or Null
    # OD_df = OD_df.drop(OD_df[(OD_df.origin_score < 1) | (OD_df.origin_score.isna())].index)
    # OD_df = OD_df.drop(OD_df[(OD_df.destination_score < 1) | (OD_df.destination_score.isna())].index)

    print(
        f"A total of {dropped_origins:n} origins and {dropped_destinations:n} destinations were invalid and not included in the analysis.")

    return OD_df

def map_input_stations(OD_df, upgrade_list):
    """ This method takes the OD Matrix generated previously and the upgrade_list where the user
      specifies the new categories for all stations in the country under each scenario. This list must contain all the
      stations in the network, not just those that are upgraded.
    ...

    :param OD_df: OD Matrix created by the input_OD_Matrix method

    :param upgrade_list: This list contains all stations and station categories in the country under each scenario.
    The list is generated from the input_df.
    """
    # A copy of the OD matrix is created before upgrading any values
    New_ODMatrix = OD_df.copy()

    # Upgrade routine for the OD Matrix. The loop is done as many times as the number of elements in the upgrade_list
    # The upgrade list only contains stations that are upgraded ignoring the stations that remain unchanged.
    for tlc in upgrade_list.TLC:
        if str(tlc) == 'nan':
            continue
        # Update category. It is necessary to use the .item() method to the Series ,
        new_category = upgrade_list.loc[upgrade_list.TLC == tlc, 'New_Category'].item()

        # Update origin category
        New_ODMatrix.loc[New_ODMatrix.OriginTLC == str(tlc), 'AfAOrigin'] = new_category
        # Update destination category
        New_ODMatrix.loc[New_ODMatrix.DestinationTLC == str(tlc), 'AfADest'] = new_category

    # After the loop is competed, we need to use the map functions again to refresh the scores
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

    # Drop unnecessary columns
    New_ODMatrix.drop(axis=1, columns=['origin_score', 'destination_score', 'jny_score'], inplace=True)

    # Create pivot table for the final output
    base_pivot = OD_df.pivot_table(index='AfAOrigin', columns='AfADest', values='Total_Journeys', aggfunc=np.sum)
    pivot = New_ODMatrix.pivot_table(index='AfAOrigin', columns='AfADest', values='Total_Journeys', aggfunc=np.sum)
    #pivot by region ...
    #tester
    pivot['Region'] ='National'

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

def regional_pivots(New_ODMatrix):
    """ This method creates a set of OD Matrices for each region using the .pivot_table method. This pivot uses the
     OD matrix as an input and the step-free categories as row and column headers. The output is a dataframe containing
     all regions.
    ...

    :param New_ODMatrix: Dataframe containing the national Origin-Destination matrix for Great Britain. The format of
    this matrix requires stations
    """

    list_of_pivots = pd.DataFrame()
    for region in New_ODMatrix['Region'].unique():

        if str(region) == '0':
            continue
        if str(region) == 'nan':
            continue
        else:
            temp_frame = New_ODMatrix.loc[New_ODMatrix['Region'] == region]
            pivot = temp_frame.pivot_table(index='AfAOrigin', columns='AfADest', values='Total_Journeys', aggfunc=np.sum)
            pivot['Region'] = region
            pivot['Scenario'] = scenario
            list_of_pivots = list_of_pivots.append(pivot)
    return list_of_pivots

def regional_kpi(New_ODMatrix, scenario_desc):
    """ This method creates a set of key performance indicators from the OD Matrix for each scenario and region.
    ...

        :param New_ODMatrix: Dataframe containing the national Origin-Destination matrix for Great Britain. The format of
        this matrix requires stations
        :param scenario_desc: scenario description provided by the user. Taken from the input template
    """

    list_of_kpi = pd.DataFrame()
    scen_name = scenario_desc.columns[0]
    for region in New_ODMatrix['Region'].unique():

        if str(region) == '0':
            continue
        if str(region) == 'nan':
            continue
        else:
            temp_frame = New_ODMatrix.loc[New_ODMatrix['Region'] == region]
            Cat_A_jnys = temp_frame.Total_Journeys.loc[(temp_frame['jny_category'] == "A")].sum()
            Cat_B1_jnys = temp_frame.Total_Journeys.loc[(temp_frame['jny_category'] == "B1")].sum()
            step_free_jnys_pctg = (Cat_A_jnys + Cat_B1_jnys) / temp_frame.Total_Journeys.sum()

            B3_stns = temp_frame.Total_Journeys.loc[(temp_frame['jny_category'] == "B3")].sum()
            C_stns = temp_frame.Total_Journeys.loc[(temp_frame['jny_category'] == "C")].sum()
            B3_or_C_stns_pctg = (B3_stns + C_stns) / temp_frame.Total_Journeys.sum()
            my_dict = {'scenario_name': scen_name, 'scenario_desc': scenario_desc.loc['Description'],
                       'scenario_notes': scenario_desc.loc['Notes'], 'step_free_journeys_%': step_free_jnys_pctg,
                       'B3_or_C_stations_%': B3_or_C_stns_pctg, 'Region': region}
            kpi_df = pd.DataFrame(data=my_dict)
            list_of_kpi = list_of_kpi.append(kpi_df)

    # National results
    Cat_A_jnys = New_ODMatrix.Total_Journeys.loc[(New_ODMatrix['jny_category'] == "A")].sum()
    Cat_B1_jnys = New_ODMatrix.Total_Journeys.loc[(New_ODMatrix['jny_category'] == "B1")].sum()
    step_free_jnys_pctg = (Cat_A_jnys + Cat_B1_jnys) / New_ODMatrix.Total_Journeys.sum()
    B3_stns = New_ODMatrix.Total_Journeys.loc[(New_ODMatrix['jny_category'] == "B3")].sum()
    C_stns = New_ODMatrix.Total_Journeys.loc[(New_ODMatrix['jny_category'] == "C")].sum()
    B3_or_C_stns_pctg = (B3_stns + C_stns) / New_ODMatrix.Total_Journeys.sum()
    my_dict = {'scenario_name': scen_name, 'scenario_desc': scenario_desc.loc['Description'],
               'scenario_notes': scenario_desc.loc['Notes'], 'step_free_journeys_%': step_free_jnys_pctg,
               'B3_or_C_stations_%': B3_or_C_stns_pctg, 'Region': 'All'}
    kpi_df = pd.DataFrame(data=my_dict)
    list_of_kpi = list_of_kpi.append(kpi_df)

    return list_of_kpi

def into_stepfree_spreadsheet(grouped_origin_df, grouped_destination_df, path_of_spreadsh, scenario_desc, st_cat_df, list_of_pivots):
    """ This method writes the results into the cloned step-free spreadsheet. The original file will never be edited.
    ...

    :param grouped_origin_df: This dataframe contains the number of trips from every station in the country. Trips from
    step-free stations (categories A and B1) are blanked so the final dataframe only has non step-free origins.
    :param grouped_destination_df: This dataframe contains the number of trips to every station in the country. Trips
    to step-free stations (categories A and B1) are blanked so the final dataframe only has non step-free destinations.

    :param path_of_spreadsh: this is the clone of the step-free scoring spreadsheet created for each scenario

    :param scenario_desc: scenario description provided by the user. Taken from the input template

    :param st_cat_df: contains the updated station cateogries, the spreadsheet will calculate all the relevant metrics

    :param list_of_pivots: dataframe generated by the regional_pivots method. This will be written into the pivot tab
    of the spreadsheet.
    """

    target = original.replace('.xlsx', '_') +str(scenario)+'.xlsx'

    #copying the path to spreadsheet file as the target scenario
    shutil.copyfile(path_of_spreadsh, target)

    #reading from the spreadsheet
    new_st_cat_df = st_cat_df.copy()

    # Next steps upgrading the data in the clone to this current scenario:
    # 1. find the stations to be upgraded.
    for code in upgrade_list.TLC:
        if str(code) == 'nan':
            continue

        #if the current code is in the station category spreadsheet then upgrade st_cat_df with new category
        if code in st_cat_df['CRS_Code'].values:
            new_category = upgrade_list.loc[upgrade_list.TLC == code, 'New_Category'].item()
            new_st_cat_df.loc[new_st_cat_df.CRS_Code == str(code), 'Including CP6 AfA'] = new_category
        else:
            continue

    # Create the kpi's that will be saved in the spreadsheet

    kpi_df = regional_kpi(New_ODMatrix, scenario_desc)

    #export back to Excel
    with pd.ExcelWriter(target, mode="a", engine="openpyxl", if_sheet_exists='replace') as writer:

        new_st_cat_df.to_excel(writer, sheet_name="St_Cat", index=False)
        kpi_df.to_excel(writer, sheet_name="KPI_Py", index=False)
        grouped_origin_df.to_excel(writer, sheet_name="Inaccessible O Accessi D")
        grouped_destination_df.to_excel(writer, sheet_name="Accessible O Inaccessi D")
        list_of_pivots.to_excel(writer, sheet_name="Pivot")

    # Open the Excel file so the formulas are calculated
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.DisplayAlerts = False
    workbook = excel.Workbooks.Open(target)
    # this must be the absolute path (r'C:/abc/def/ghi')
    # workbook.RefreshAll()
    workbook.Save()
    workbook.Close()
    excel.Quit()

    #done


def output_to_log(input_df, scenario):
    """ This method writes the list of upgraded stations under each scenario into a txt file for reference.
    ...

    :param input_df: This dataframe contains.
    :param scenario: scenario name.

    """
    now = datetime.now()
    dt_string = now.strftime("%d/%m/%Y %H:%M:%S")

    with open(output_log_path, "a+") as file_object:
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

def make_kepler_input(target, scenario):
    """ This method creates the csv file that will be passed to Kepler to build maps of the output.
    https://kepler.gl/demo
    ...

    :param scenario: scenario name.
    :param path_of_spreadsh: this is the clone of the step-free scoring spreadsheet created for each scenario

    """
    # Grabbing the all stations tab from the selected scenario
    final_df = pd.read_excel(target, sheet_name="All Stations", engine='openpyxl')
    final_df.columns = [c.replace(' ', '_') for c in final_df.columns]

    #naming this dataframe kepler as it will serve as the stations.csv file for kepler
    kepler = pd.read_excel(target, sheet_name="Coordinates", header=2, usecols="B:F", engine='openpyxl')

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
            kepler.loc[kepler.CRSCode == str(code), 'Isolation Score']  = final_df.loc[final_df.Unique_Code == code, 'Mobility_and_Isolation_Matrix_Outcome'].item()
            kepler.loc[kepler.CRSCode == str(code), 'Socioeconomic Score'] = final_df.loc[final_df.Unique_Code == code, 'Socioecon-Local_Impact_Matrix_Outcome'].item()
            kepler.loc[kepler.CRSCode == str(code), 'Final Matrix Outcome']  = final_df.loc[final_df.Unique_Code == code, 'Final_Outcome'].item()
            #kepler.loc[kepler.CRSCode == str(code), 'DfT Category A to F']  = get_DfT_Num_Cat(final_df.loc[final_df.Unique_Code == code, 'Dft_Category'].item())

    #Exporting the file as a csv
    csv_outpath = kepler_path + '\Stations_sc_' +str(scenario) + '.csv'
    kepler.to_csv(csv_outpath)

def scenario_input():

    input_df = pd.read_excel(input_path, sheet_name="Diffs", engine='openpyxl', index_col=0)
    input_df.drop(['Station Name (MOIRA Name)', 'CP5', 'CP6'], axis=1, inplace=True)
    input_df.replace(to_replace="ignore", value=np.NaN, inplace=True)
    all_scen_df = pd.read_excel(input_path, sheet_name="Scenario Master", engine='openpyxl', index_col=0)

    return all_scen_df, input_df


#Pseudo-Main

shutil.copyfile(original, clone)

path_of_spreadsh = clone

all_scen_df, input_df =scenario_input()

st_cat_df = pd.read_excel(path_of_spreadsh, sheet_name="St_Cat", engine='openpyxl')
st_cat_df.rename(columns={'CRS Code': 'CRS_Code', 'Station Name (MOIRA Name)': 'Station_Name'}, inplace=True)
st_cat_df = st_cat_df.loc[:,~st_cat_df.columns.duplicated()].copy()

OD_df = input_OD_Matrix(st_cat_df)

# Base scenario run
scenario = 'CP6'
scenario_desc = pd.DataFrame(all_scen_df.loc[scenario][['Description', 'Notes']])
print(f'{scenario} Started')
target = original.replace('.xlsx', '_') + str(scenario) + '.xlsx'

shutil.copyfile(path_of_spreadsh, target)
kpi_df = regional_kpi(OD_df, scenario_desc)
list_of_pivots = regional_pivots(OD_df)

with pd.ExcelWriter(target, mode="a", engine="openpyxl", if_sheet_exists='replace') as writer:

    kpi_df.to_excel(writer, sheet_name="KPI_Py", index=False)
    list_of_pivots.to_excel(writer, sheet_name="Pivot")

# Open the Excel file so the formulas are calculated
excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.DisplayAlerts = False
workbook = excel.Workbooks.Open(target)
# this must be the absolute path (r'C:/abc/def/ghi')
# workbook.RefreshAll()
workbook.Save()
# workbook.Save() will update the formulas but it brings a dialog box which stops the execution
# of the code until the user selects a level of sensitivity (Public, General, Confidential, etc)
workbook.Close()
excel.Quit()

make_kepler_input(target, scenario)

print(f"Base scenario run successfully.")


# Run all scenarios
for scenario in input_df.columns:

    if input_df[scenario].isnull().all():
        print(f'Empty scenario. {scenario} Skipped')
        continue
    if (input_df[scenario] == 0).all():
        print(f'This scenario has no changes. {scenario} Skipped')
        continue
    else:
        # clones spreadsheet as to not affect the original when writing to the sheet
        print(f'{scenario} Started')
        shutil.copyfile(original, clone)

        scenario_desc = pd.DataFrame(all_scen_df.loc[scenario][['Description', 'Notes']])
        path_of_spreadsh = clone

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
        list_of_pivots = regional_pivots(New_ODMatrix)

        ##
        output_to_log(upgrade_list, str(scenario))
        #
        into_stepfree_spreadsheet(grouped_origin_df, grouped_destination_df, path_of_spreadsh, scenario_desc, st_cat_df, list_of_pivots)
        #
        make_kepler_input(target, scenario)

        print(f"Scenario {scenario} run successfully. {len(upgrade_list)} stations were upgraded")

os.remove(clone)
print(f"Process finished successfully")
