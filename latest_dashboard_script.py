'''
Authors: Kharesa-Kesa and Jose
This script will attempt to replicate the actions of the spreadsheet generating the all_stations tab
'''

# from asyncio.unix_events import _UnixSelectorEventLoop
# from email import header
from operator import index
from matplotlib.pyplot import polar
import pandas as pd, numpy as np, glob, ast, openpyxl, shutil, pyodbc, random, datetime
from openpyxl import Workbook
from openpyxl import load_workbook
from datetime import datetime


def get_orr_step_free_category(search_value):
    orr = {'B': 'Bottom', 'B2': 'Bottom', 'B3': 'Middle', 'C': 'Top'}

    return orr[search_value]


def get_connectivity_journeys_matrix(search_value):
    e = {'TopTop': 1, 'TopMiddle': 2, 'TopBottom': 3, 'MiddleTop': 1, 'MiddleMiddle': 2, 'MiddleBottom': 3,
         'BottomTop': 2, 'BottomMiddle': 3, 'BottomBottom': 5}

    return e[search_value]


def get_updated_stations():
    input_path = r'https://networkrail-my.sharepoint.com/:x:/r/personal/kadams5_networkrail_co_uk/Documents/Arup/National%20Rail%20Accessibility%20Strategy/04%20Working%20folder/05%20Dev/Input%20template.csv'
    input_df = pd.read_csv(input_path)

    return input_df


def get_mobility_isolation_matrix(search_value):
    e = {'TopTop': 1, 'TopMiddle': 1, 'TopBottom': 2, 'MiddleTop': 1, 'MiddleMiddle': 2, 'MiddleBottom': 3,
         'BottomTop': 2, 'BottomMiddle': 3, 'BottomBottom': 3}

    return e[search_value]


def get_list_col():
    list_cols = ['Station_Name', 'Unique_Code', 'Station_Facility_Owner',
                 'Network_Rail_Region', 'ORR_Step_Free_Category', 'DfT_Category',
                 'Inaccessible_(1_if_not_Step_Free_Cat._A_or_B1)',
                 '2018/2019_ORR_Total_Entries/Exits_and_Interchange',
                 '2019_Journeys_to_an_accessible_destination',
                 '2019_Journeys_from_an_accessible_origin',
                 '2019_Total_Unlocked_Journeys', '2019_Potential_Unlocked_Rank',
                 '2019_Unlocked_Journeys_Percentile',
                 '2019_Unlocked_Journeys_Matrix_Outcome',
                 '2019_Connectivity_(count_of_stations_directly_served)',
                 '2019_Connectivity_Rank', '2019_Connectivity_Percentile',
                 '2019_Connectivity_Matrix_Outcome',
                 'Connectivity_and_Journeys_Matrix_Outcome',
                 'Connectivity_and_Journeys_Matrix_Outcome.1', 'Mobility_Score',
                 'Isolation_(1_if_no_Cat_A_in_20_mins_drive_isochrone)',
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

    one = ['Station_Name', 'Unique_Code', 'Station_Facility_Owner',
           'Network_Rail_Region', 'ORR_Step_Free_Category', 'DfT_Category',
           'Inaccessible_(1_if_not_Step_Free_Cat._A_or_B1)',
           '2018/2019_ORR_Total_Entries/Exits_and_Interchange',
           '2019_Journeys_to_an_accessible_destination',
           '2019_Journeys_from_an_accessible_origin',
           '2019_Total_Unlocked_Journeys', '2019_Potential_Unlocked_Rank',
           '2019_Unlocked_Journeys_Percentile',
           '2019_Unlocked_Journeys_Matrix_Outcome',
           '2019_Connectivity_(count_of_stations_directly_served)',
           '2019_Connectivity_Rank', '2019_Connectivity_Percentile']

    two = ['Inaccessible_(1_if_not_Step_Free_Cat._A_or_B1)', '2019_Connectivity_Matrix_Outcome',
           'Connectivity_and_Journeys_Matrix_Outcome',
           'Connectivity_and_Journeys_Matrix_Outcome.1', 'Mobility_Score',
           'Isolation_(1_if_no_Cat_A_in_20_mins_drive_isochrone)',
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


def input_OD_Matrix():
    # inputs
    # temporary variable to direct the code, when the access database is updated then we can remove
    reading_from_access = True

    if reading_from_access:

        # connect to the access database
        conn = pyodbc.connect(
            r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=https://networkrail-my.sharepoint.com/:u:/r/personal/kadams5_networkrail_co_uk/Documents/Arup/National%20Rail%20Accessibility%20Strategy/02%20Step%20Free%20Data%20Analysis/05%20MOIRA/MOIRAOD.accdb;')
        '''
        cursor = conn.cursor()
        query = cursor.execute('select * from ODMatrix')
        for row in cursor.fetchall():
            print (row)
        '''
        query = 'select * from ODMatrixAfA'
        OD_df = pd.read_sql(query, conn)

    else:
        # reading the data in from CSVs manaually from local storage
        vs1 = pd.read_csv(
            r'C:\Users\Kharesa-Kesa.Spencer\OneDrive - Arup\Projects\Network Rail Accessibility case\matrices\Vector save.csv')
        vs2 = pd.read_csv(
            r'C:\Users\Kharesa-Kesa.Spencer\OneDrive - Arup\Projects\Network Rail Accessibility case\matrices\Vector save 2.csv')
        frames = [vs1, vs2]
        OD_df = pd.concat(frames, ignore_index=True)

    return OD_df


def map_input_stations(OD_df, base_df):
    # Add numerical score based on a dictionary
    my_dict = {'0': 0, 'A': 1, 'B': 2, 'B1': 3, 'B2': 4, 'B3': 5, 'C': 6, 'Null': -1}
    origin_score = OD_df.AfAOrigin.map(my_dict)
    destination_score = OD_df.AfADest.map(my_dict)
    OD_df['origin_score'] = origin_score
    OD_df['destination_score'] = destination_score

    # Remove OD pairs where the score is 0 or Null. Save the number of removed records for traceability
    dropped_origins = len(OD_df[(OD_df.origin_score < 1) | (OD_df.origin_score.isna())])
    dropped_destinations = len(OD_df[(OD_df.destination_score < 1) | (OD_df.destination_score.isna())])
    OD_df = OD_df.drop(OD_df[(OD_df.origin_score < 1) | (OD_df.origin_score.isna())].index)
    OD_df = OD_df.drop(OD_df[(OD_df.destination_score < 1) | (OD_df.destination_score.isna())].index)

    print(f"A total of {dropped_origins:n} origins and {dropped_destinations:n} destinations were invalid and not included in the analysis.")

    # Categorize the journeys by looking at the maximum numerical score of the origin and destination stations
    jny_score = np.maximum(OD_df.origin_score, OD_df.destination_score)

    # Add the list as a new column to the existing dataframe
    OD_df['jny_score'] = jny_score
    my_dict_2 = {1: 'A', 2: 'B', 3: 'B1', 4: 'B2', 5: 'B3', 6: 'C'}
    jny_category = OD_df.jny_score.map(my_dict_2)
    OD_df['jny_category'] = jny_category
    OD_df['Total_Journeys'] = OD_df['STDJOURNEYS'] + OD_df['1stJOURNEYS']
    # print(base_df.tail(10))

    ## TEST
    OD_df.isna().sum()

    # base_df.info()
    #v1 = OD_df.groupby("jny_category").Total_Journeys.sum()
    # print(v1)
    # print(type(v1))
    # print(base_df.tail(10))

    # this method does all the calculating the new categories

    # Import stations to be upgraded
    upgrade_list = get_updated_stations()

    # In anticipation of the new score the columm is set to None
    base_df['Inaccessible_(1_if_not_Step_Free_Cat._A_or_B1)'] = None

    for tlc in upgrade_list.TLC:
        if str(tlc) == 'nan':
            continue
        # Update category. It is necessary to use the .item() method to the Series ,
        new_category = upgrade_list.loc[upgrade_list.TLC == tlc, 'New_Category'].item()

        # Update origin category
        OD_df.loc[OD_df.OriginTLC == str(tlc), 'AfAOrigin'] = new_category
        # Update destination category
        OD_df.loc[OD_df.DestinationTLC == str(tlc), 'AfADest'] = new_category

        if tlc in base_df.values:
            # Update ORR_Step_Free_Category
            base_df.loc[base_df.Unique_Code == str(tlc), 'ORR_Step_Free_Category'] = new_category
            # setting their indicator to 0 as they are now accessible
            base_df.loc[base_df.Unique_Code == str(tlc), 'Inaccessible_(1_if_not_Step_Free_Cat._A_or_B1)'] = 0

    # Update origin score
    OD_df.origin_score = OD_df.AfAOrigin.map(my_dict)
    # Update destination score
    OD_df.destination_score = OD_df.AfADest.map(my_dict)
    # Update journey score
    OD_df.jny_score = np.maximum(OD_df.origin_score, OD_df.destination_score)
    # Update journey category
    OD_df.jny_category = OD_df.jny_score.map(my_dict_2)
    # Concat the 2 categories together
    OD_df['concat_categories'] = OD_df.AfAOrigin + OD_df.AfADest

    New_ODMatrix = OD_df
    New_ODMatrix.drop(axis=1,columns=['origin_score', 'destination_score', 'jny_score', 'jny_category'], inplace=True)

    # dataframes where only the origin or the destination are accessible
    OD_df_ass_origin = OD_df.loc[(OD_df.AfAOrigin == 'A') | (OD_df.AfAOrigin == 'B1')]
    OD_df_ass_destination = OD_df.loc[(OD_df.AfADest == 'A') | (OD_df.AfADest == 'B1')]

    grouped_origin_df = (OD_df_ass_origin.groupby(["OriginTLC", "AfAOrigin"])["Total_Journeys"].sum()).to_frame()
    grouped_origin_df.reset_index(inplace=True)

    grouped_destination_df = (
        OD_df_ass_destination.groupby(["DestinationTLC", "AfADest"])["Total_Journeys"].sum()).to_frame()
    grouped_destination_df.reset_index(inplace=True)


    # Placing the total journey grouped values into
    for code in base_df.Unique_Code:
        if str(tlc) == 'nan':
            continue

        if str(code) in grouped_origin_df.values:
            total_jo = grouped_origin_df.loc[grouped_origin_df.OriginTLC == str(code), 'Total_Journeys'].item()
            base_df.loc[base_df.Unique_Code == str(code), '2019_Journeys_from_an_accessible_origin'] = total_jo

        if str(code) in grouped_destination_df.values:
            total_jd = grouped_destination_df.loc[
                grouped_destination_df.DestinationTLC == str(code), 'Total_Journeys'].item()
            base_df.loc[base_df.Unique_Code == str(code), '2019_Journeys_to_an_accessible_destination'] = total_jd

    return base_df


def get_new_categories_set_jrnys(base_df):
    # set the OD matrix, either from the access database or from the csv
    OD_df = input_OD_Matrix()

    base_df = map_input_stations(OD_df, base_df)

    # so now the base df and the od df both have updated station categories from the input template
    # next is adding the column codes and the journey stats, mapping from the OD_df to the base df
    # if a station is accessible they're changed to 0 and if the column is empty its set to NaN
    # Did it manually due to error
    base_df.loc[base_df.ORR_Step_Free_Category == 'A', 'Inaccessible_(1_if_not_Step_Free_Cat._A_or_B1)'] = 0
    base_df.loc[base_df.ORR_Step_Free_Category == 'B1', 'Inaccessible_(1_if_not_Step_Free_Cat._A_or_B1)'] = 0
    base_df.loc[base_df.ORR_Step_Free_Category == 'B', 'Inaccessible_(1_if_not_Step_Free_Cat._A_or_B1)'] = 1
    base_df.loc[base_df.ORR_Step_Free_Category == 'B2', 'Inaccessible_(1_if_not_Step_Free_Cat._A_or_B1)'] = 1
    base_df.loc[base_df.ORR_Step_Free_Category == 'B3', 'Inaccessible_(1_if_not_Step_Free_Cat._A_or_B1)'] = 1
    base_df.loc[base_df.ORR_Step_Free_Category == 'C', 'Inaccessible_(1_if_not_Step_Free_Cat._A_or_B1)'] = 1

    base_df.loc[base_df.ORR_Step_Free_Category.isna(), 'Inaccessible_(1_if_not_Step_Free_Cat._A_or_B1)'] = np.NaN

    unlock_jrny = []
    unlocked_jrny_perc = []
    connc = []
    connc_rank = []
    connc_rank_perc = []
    connc_jrnys_matrx = []

    matrix_outcome_list = []
    matrix_outcome_list_2 = []

    for index, row in base_df.iterrows():

        if row['Inaccessible_(1_if_not_Step_Free_Cat._A_or_B1)'] == 1:

            unlock_jrny.append(
                int(row['2019_Journeys_to_an_accessible_destination'] + row['2019_Journeys_from_an_accessible_origin']))
            connc.append(int(row['2019_Connectivity_(count_of_stations_directly_served)']))

            # setting this by looping through df with if statements
        else:
            unlock_jrny.append(None)
            connc.append(None)

    base_df['2019_Total_Unlocked_Journeys'] = unlock_jrny
    unlock_jrny_s = pd.Series(unlock_jrny)
    base_df['2019_Potential_Unlocked_Rank'] = unlock_jrny_s.rank(ascending=False)
    base_df['2019_Unlocked_Journeys_Percentile'] = unlock_jrny_s.rank(ascending=False, pct=True)
    base_df['2019_Connectivity_(count_of_stations_directly_served)'] = connc
    connc_s = pd.Series(connc)
    base_df['2019_Connectivity_Rank'] = connc_s.rank(ascending=False)
    base_df['2019_Connectivity_Percentile'] = connc_s.rank(ascending=False, pct=True)

    for index, row in base_df.iterrows():
        val = row['2019_Unlocked_Journeys_Percentile']
        val_2 = row['2019_Connectivity_Percentile']

        if row['Inaccessible_(1_if_not_Step_Free_Cat._A_or_B1)'] == 1:

            if val < 0.33:
                matrix_outcome_list.append('Top')

            elif val < 0.66:
                matrix_outcome_list.append('Middle')

            else:
                matrix_outcome_list.append('Bottom')

        else:
            matrix_outcome_list.append('')

        if row['Inaccessible_(1_if_not_Step_Free_Cat._A_or_B1)'] == 1:

            if val_2 < 0.33:
                matrix_outcome_list_2.append('Top')

            elif val_2 < 0.66:
                matrix_outcome_list_2.append('Middle')

            else:
                matrix_outcome_list_2.append('Bottom')

        else:
            matrix_outcome_list_2.append('')

    base_df['2019_Unlocked_Journeys_Matrix_Outcome'] = matrix_outcome_list
    base_df['2019_Connectivity_Matrix_Outcome'] = matrix_outcome_list_2
    base_df['Connectivity_and_Journeys_Matrix_Outcome'] = base_df['2019_Unlocked_Journeys_Matrix_Outcome'] + base_df[
        '2019_Connectivity_Matrix_Outcome']

    # Matrix
    for index, row in base_df.iterrows():
        mob = str(row['Connectivity_and_Journeys_Matrix_Outcome'])

        if mob == 'nan' or mob == '':
            continue
        else:
            base_df.loc[base_df[
                            'Connectivity_and_Journeys_Matrix_Outcome'] == mob, 'Connectivity_and_Journeys_Matrix_Outcome.1'] = get_connectivity_journeys_matrix(
                mob)

    return base_df


def set_mobility_isolation_score(updated_cats_and_jrnys, alt_any_df):
    # COLUMN V

    # Matrix
    for index, row in updated_cats_and_jrnys.iterrows():

        if row['Inaccessible_(1_if_not_Step_Free_Cat._A_or_B1)'] == 1:
            osfc = str(row['ORR_Step_Free_Category'])

            updated_cats_and_jrnys.loc[updated_cats_and_jrnys['Unique_Code'] == str(
                row['Unique_Code']), 'Mobility_Score'] = get_orr_step_free_category(osfc)

        else:
            updated_cats_and_jrnys.loc[
                updated_cats_and_jrnys['Unique_Code'] == str(row['Unique_Code']), 'Mobility_Score'] = None

    # getting the input stations
    input_df = get_updated_stations()
    list_of_changed_stns = list(input_df['TLC'])

    # COLUMN Z
    # creating a blank dataframe with the same columns as the alt_any sheet table
    stns_df = pd.DataFrame(columns=['Station_Code', 'Station_Name', 'Region', 'ID2', 'Station__2'])

    # looping through the list of inputted stations
    # for every station thats in the in the dataframe
    # check if this is the first time the dataframe is being appended (through length) if not then
    # append all the rows where the station code equals the one being looped through
    for station in list_of_changed_stns:

        if station in alt_any_df.values:

            if len(stns_df) == 0:
                stns_df = alt_any_df[alt_any_df.Station_Code == station]

            else:
                temp_df = alt_any_df[alt_any_df.Station_Code == station]
                stns_df = pd.concat([stns_df, temp_df], axis=0, )

    # a list of all the station codes of the stations within 20 of all the newly upgraded stations (using set to remove duplicates)
    target_stns = list(set(stns_df.ID2))

    # now we pivot over top the all station sheet to assign bottom to all these stations

    for target in target_stns:

        if target in updated_cats_and_jrnys.values:
            updated_cats_and_jrnys.loc[(updated_cats_and_jrnys.Unique_Code == target) & (updated_cats_and_jrnys[
                                                                                             'Inaccessible_(1_if_not_Step_Free_Cat._A_or_B1)'] == 1), 'Revisited_Isolation_score'] = 'Bottom'

    # COLUMN Y

    updated_cats_and_jrnys['Mobility/Isolation'] = updated_cats_and_jrnys['Original_Isolation_Score'] + \
                                                   updated_cats_and_jrnys['Revisited_Isolation_score']

    # Matrix
    for index, row in updated_cats_and_jrnys.iterrows():
        mob = str(row['Mobility/Isolation'])

        if mob == 'nan':
            continue
        else:
            updated_cats_and_jrnys.loc[updated_cats_and_jrnys[
                                           'Mobility/Isolation'] == mob, 'Isolation_and_Current_Access_Matrix_Outcome'] = get_mobility_isolation_matrix(
                mob)

    # dropping final duplicates before returning
    updated_cats_and_jrnys = updated_cats_and_jrnys.loc[:, ~updated_cats_and_jrnys.columns.duplicated()].copy()

    return updated_cats_and_jrnys


def blanking_rows(updated_mobility_and_isolation):
    list_of_cols = get_list_col()

    for index, row in updated_mobility_and_isolation.iterrows():

        for col in list_of_cols:

            if row['Inaccessible_(1_if_not_Step_Free_Cat._A_or_B1)'] == 0:
                updated_mobility_and_isolation[col] = None

    return updated_mobility_and_isolation


path_of_spreadsh = r'https://networkrail-my.sharepoint.com/:x:/r/personal/kadams5_networkrail_co_uk/Documents/Arup/National%20Rail%20Accessibility%20Strategy/04%20Working%20folder/05%20Dev/Step%20Free%20Scoring_JDL_v3.00.xlsx'
base_df = pd.read_excel(path_of_spreadsh, sheet_name="All Stations", header=2, usecols="B:AS", engine='openpyxl')
alt_any = pd.read_excel(path_of_spreadsh, sheet_name="Alt_Any_20", header=4, usecols="B:F", engine='openpyxl')

base_df.columns = [c.replace(' ', '_') for c in base_df.columns]
alt_any.columns = [c.replace(' ', '_') for c in alt_any.columns]

updated_cats_and_jrnys = get_new_categories_set_jrnys(base_df)
#
updated_mobility_and_isolation = set_mobility_isolation_score(updated_cats_and_jrnys, alt_any)
#
final_df = blanking_rows(updated_mobility_and_isolation)
