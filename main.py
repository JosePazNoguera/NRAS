"""
NRAS
Author: Jose De la Paz Noguera and Kharesa-Kesa Spencer
Date: 

Description:


This work tries to estimate the category of a journey based on the categories of the
origin and destination stations.
The category of a station can be:
    Cat     |   Description
    --------|---------------
    A       |   Step-free access available from the station to the platform (level-boarding is not within scope)
    B       |   Somewhere in between A and C. The three sub-divisions of B try to prioritise this category
    B1      |
    B2      |
    B3      |
    C       |   The station does not have step-free access facilities
    0       |   Ignored
    Null    |   Ignored


The aim of this file is to:
1. combine several csv into a single dataframe
2. classify the journeys based on the O/D station categories
"""

import pandas as pd, numpy as np, glob, pyodbc, ast

def reading_input():
    
    #temporary variable to direct the code, when the access database is updated then we can remove
    reading_from_access = False

    if reading_from_access:

        #connect to the access database
        conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\Kharesa-Kesa.Spencer\OneDrive - Arup\Projects\Network Rail Accessibility case\matrices\MOIRAOD.accdb;')
        '''
        cursor = conn.cursor()
        query = cursor.execute('select * from ODMatrix')
        for row in cursor.fetchall():
            print (row)
        '''
        query = 'select * from ODMatrix'
        my_df = pd.read_sql(query, conn)

    else:
        #reading the files in manaually from local storage
        vs1 = pd.read_csv(PATH)
        vs2 = pd.read_csv(PATH)
        frames = [vs1,vs2]
        my_df = pd.concat(frames, ignore_index=True)



    #missing values

    #list of the columns with nan values
    nan_cols = my_df.loc[:,my_df.isna().any(axis=0)]

    #if there is only one column check it isnt region, region can be ignored - metadata
    if len(nan_cols.columns) == 0:
        if nan_cols.columns[0] == 'Region':
            print('only region has empty rows, proceeding')
            
    #else drop all rows within the nan columns with empty rows        
    else:
        #if 'Total_Journeys' in nan_cols.columns:

        #first does all rows with no data
        my_df.dropna(axis=0,how='all',subset=nan_cols.columns)

        #then does all rows with no data
        my_df.dropna(axis=0,how='any',subset=nan_cols.columns)

    return my_df

def calculations(base_df):
    
    # Add numerical score based on a dictionary
    my_dict = {'0': 0,'A': 1, 'B': 2, 'B1': 3, 'B2': 4, 'B3': 5, 'C': 6, 'Null': -1}
    origin_score = base_df.Origin_Category.map(my_dict)
    destination_score = base_df.Destination_Category.map(my_dict)
    base_df['origin_score'] = origin_score
    base_df['destination_score'] = destination_score

    # Remove OD pairs where the score is 0 or Null. Save the number of removed records for traceability
    dropped_origins = base_df.origin_score[base_df.origin_score < 1].count()
    dropped_destinations = base_df.destination_score[base_df.destination_score < 1].count()
    base_df = base_df.drop(base_df[base_df.origin_score < 1].index)
    base_df = base_df.drop(base_df[base_df.destination_score < 1].index)
    # base_df.info()
    # print(base_df.dtypes)

    print(f"A total of {dropped_origins:n} origins and {dropped_destinations:n} destinations were invalid and not included in the analysis.")

    # Categorize the journeys by looking at the maximum numerical score of the origin and destination stations
    jny_score = np.maximum(base_df.origin_score, base_df.destination_score)

    # Add the list as a new column to the existing dataframe
    base_df['jny_score'] = jny_score
    my_dict_2 = {1: 'A', 2: 'B', 3: 'B1', 4: 'B2', 5: 'B3', 6: 'C'}
    jny_category = base_df.jny_score.map(my_dict_2)
    base_df['jny_category'] = jny_category
    # print(base_df.tail(10))

    #base_df.info()
    v1 = base_df.groupby("jny_category").Total_Journeys.sum()
    # print(v1)
    #print(type(v1))
    #print(base_df.tail(10))

    ### STATION UPGRADE ROUTINE

    # Import stations to be upgraded
    input_path = '#PATH OF INPUT CSV'
    upgrade_list = pd.read_csv(input_path)

    # Transform the list into a dataframe
    upgrade_list.columns = [c.replace(' ','_') for c in upgrade_list.columns]
    # print(upgrade_list.info())

    # Data columns (total 3 columns):
    #  #   Column        Non-Null Count  Dtype
    # ---  ------        --------------  -----
    #  0   Station       2 non-null      object
    #  1   TLC           5 non-null      object
    #  2   New_Category  6 non-null      object
    # dtypes: object(3)



    ### STATION UPGRADE ROUTINE

    # Import stations to be upgraded
    input_path = '#PATH OF INPUT CSV'
    upgrade_list = pd.read_csv(input_path)

    # Transform the list into a dataframe
    upgrade_list.columns = [c.replace(' ','_') for c in upgrade_list.columns]
    # print(upgrade_list.info())

    # Data columns (total 3 columns):
    #  #   Column        Non-Null Count  Dtype
    # ---  ------        --------------  -----
    #  0   Station       2 non-null      object
    #  1   TLC           5 non-null      object
    #  2   New_Category  6 non-null      object
    # dtypes: object(3)


    # Create a copy of the base data frame
    scenario_1 = base_df.copy()
    # print(type(scenario_1))

    # Next steps:
    # 1. find the stations to be upgraded.
    for tlc in upgrade_list.TLC:
        if str(tlc) == 'nan':
            continue
        # Update category. It is necessary to use the .item() method to the Series ,
        new_category = upgrade_list.loc[upgrade_list.TLC == tlc, 'New_Category'].item()

        # Update origin category
        scenario_1.loc[scenario_1.Origin_TLC == str(tlc), 'Origin_Category'] = new_category
        # Update destination category
        scenario_1.loc[scenario_1.Destination_TLC == str(tlc), 'Destination_Category'] = new_category

    # Update origin score
    scenario_1.origin_score = scenario_1.Origin_Category.map(my_dict)
    # Update destination score
    scenario_1.destination_score = scenario_1.Destination_Category.map(my_dict)
    # Update journey score
    scenario_1.jny_score = np.maximum(scenario_1.origin_score, scenario_1.destination_score)
    # Update journey category
    scenario_1.jny_category = scenario_1.jny_score.map(my_dict_2)

    return scenario_1


def main():
    #Initalising the base df through the inputs
    base_df = reading_input()
    print('Database loaded in successfully of shape: ', base_df.shape)


    # removing O-D pairs with 0 or null and categorising journeys 
    scenario_1 = calculations()

