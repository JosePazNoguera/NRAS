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


def get_new_categories_set_jrnys(base_df):


    #inputs
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
        OD_df = pd.read_sql(query, conn)

    else:
        #reading the data in from CSVs manaually from local storage
        vs1 = pd.read_csv('/Users/kharesa-kesa.spencer/Library/CloudStorage/OneDrive-Arup/Projects/Network Rail Accessibility case/matrices/Vector save.csv')
        vs2 = pd.read_csv('/Users/kharesa-kesa.spencer/Library/CloudStorage/OneDrive-Arup/Projects/Network Rail Accessibility case/matrices/Vector save 2.csv')
        frames = [vs1,vs2]
        OD_df = pd.concat(frames, ignore_index=True)


    # Add numerical score based on a dictionary
    my_dict = {'0': 0,'A': 1, 'B': 2, 'B1': 3, 'B2': 4, 'B3': 5, 'C': 6, 'Null': -1}
    origin_score = OD_df.Origin_Category.map(my_dict)
    destination_score = OD_df.Destination_Category.map(my_dict)
    OD_df['origin_score'] = origin_score
    OD_df['destination_score'] = destination_score

    # Remove OD pairs where the score is 0 or Null. Save the number of removed records for traceability
    dropped_origins = OD_df.origin_score[OD_df.origin_score < 1].count()
    dropped_destinations = OD_df.destination_score[OD_df.destination_score < 1].count()
    OD_df = OD_df.drop(OD_df[OD_df.origin_score < 1].index)
    OD_df = OD_df.drop(OD_df[OD_df.destination_score < 1].index)

    print(f"A total of {dropped_origins:n} origins and {dropped_destinations:n} destinations were invalid and not included in the analysis.")

    # Categorize the journeys by looking at the maximum numerical score of the origin and destination stations
    jny_score = np.maximum(OD_df.origin_score, OD_df.destination_score)

    # Add the list as a new column to the existing dataframe
    OD_df['jny_score'] = jny_score
    my_dict_2 = {1: 'A', 2: 'B', 3: 'B1', 4: 'B2', 5: 'B3', 6: 'C'}
    jny_category = OD_df.jny_score.map(my_dict_2)
    OD_df['jny_category'] = jny_category
    # print(base_df.tail(10))

    #base_df.info()
    v1 = OD_df.groupby("jny_category").Total_Journeys.sum()
    # print(v1)
    #print(type(v1))
    #print(base_df.tail(10))



    #this method does all the calculating the new categories

    # Import stations to be upgraded
    input_path = '/Users/kharesa-kesa.spencer/Library/CloudStorage/OneDrive-Arup/Projects/Network Rail Accessibility case/matrices/Input template.csv'
    upgrade_list = pd.read_csv(input_path)



    for tlc in upgrade_list.TLC:
        if str(tlc) == 'nan':
            continue
        # Update category. It is necessary to use the .item() method to the Series ,
        new_category = upgrade_list.loc[upgrade_list.TLC == tlc, 'New_Category'].item()

        # Update origin category
        OD_df.loc[OD_df.Origin_TLC == str(tlc), 'Origin_Category'] = new_category
        # Update destination category
        OD_df.loc[OD_df.Destination_TLC == str(tlc), 'Destination_Category'] = new_category

        if tlc in base_df.values:
            # Update ORR_Step_Free_Category
            base_df.loc[base_df.Unique_Code == str(tlc), 'ORR_Step_Free_Category'] = new_category

    # Update origin score
    OD_df.origin_score = OD_df.Origin_Category.map(my_dict)
    # Update destination score
    OD_df.destination_score = OD_df.Destination_Category.map(my_dict)
    # Update journey score
    OD_df.jny_score = np.maximum(OD_df.origin_score, OD_df.destination_score)
    # Update journey category
    OD_df.jny_category = OD_df.jny_score.map(my_dict_2)
    # Concat the 2 categories together
    OD_df['concat_categories'] = OD_df.Origin_Category + OD_df.Destination_Category
    #dataframes where only the origin or the destination are accessible 
    OD_df_ass_origin = OD_df.loc[(OD_df.Origin_Category =='A') | (OD_df.Origin_Category =='B1')]
    OD_df_ass_destination = OD_df.loc[(OD_df.Destination_Category  =='A') | (OD_df.Destination_Category  =='B1')]


    grouped_origin_df = (OD_df_ass_origin.groupby(["Origin_TLC","Origin_Category"])["Total_Journeys"].sum()).to_frame()
    grouped_origin_df.reset_index(inplace=True)

    grouped_destination_df= (OD_df_ass_destination.groupby(["Destination_TLC","Destination_Category"])["Total_Journeys"].sum()).to_frame()
    grouped_destination_df.reset_index(inplace=True)


    #Placing the total journey grouped values into 
    for code in base_df.Unique_Code:
        if str(tlc) == 'nan':
            continue

        if str(code) in grouped_origin_df.values:

            total_jo = grouped_origin_df.loc[grouped_origin_df.Origin_TLC == str(code), 'Total_Journeys'].item()
            base_df.loc[base_df.Unique_Code == str(code), '2019_Journeys_from_an_accessible_origin'] = total_jo

        if str(code) in grouped_destination_df.values:
            total_jd = grouped_destination_df.loc[grouped_destination_df.Destination_TLC == str(code), 'Total_Journeys'].item()
            base_df.loc[base_df.Unique_Code == str(code), '2019_Journeys_to_an_accessible_destination'] = total_jd


    #so now the base df and the od df both have updated station categories from the input template
    #next is adding the column codes and the journey stats, mapping from the OD_df to the base df
    #all stations are set to 1, then if they're accessible they're changed to 0 and if the column is empty its set to NaN
    base_df['Inaccessible_(1_if_not_Step_Free_Cat._A)'] = 1
    base_df.loc[(base_df.ORR_Step_Free_Category == 'A') | (base_df.ORR_Step_Free_Category == 'B1'), 'Inaccessible_(1_if_not_Step_Free_Cat._A)'] = 0
    base_df.loc[base_df.ORR_Step_Free_Category.isna(), 'Inaccessible_(1_if_not_Step_Free_Cat._A)'] = np.NaN

    #totals
    base_df['2019_Total_Unlocked_Journeys'] = base_df['2019_Journeys_to_an_accessible_destination'] + base_df['2019_Journeys_from_an_accessible_origin']
    base_df['2019_Potential_Unlocked_Rank'] = base_df['2019_Total_Unlocked_Journeys'].rank(ascending=False)
    base_df['2019_Unlocked_Journeys_Percentile'] = base_df['2019_Total_Unlocked_Journeys'].rank(ascending=False, pct=True)


    #setting this by looping through df with if statements
    matrix_outcome_list = []

    for val in base_df['Inaccessible_(1_if_not_Step_Free_Cat._A)']:

        if  val == 1:

            if val< 0.33:
                matrix_outcome_list.append('Top')

            elif val < 0.66:
                matrix_outcome_list.append('Middle')
            
            else:
                matrix_outcome_list.append('Bottom')
        
        else:
            matrix_outcome_list.append('')

    base_df['2019_Unlocked_Journeys_Matrix_Outcome'] = matrix_outcome_list
    base_df['2019_Connectivity_Rank'] = base_df['2019_Connectivity_(count_of_stations_directly_served)'].rank(ascending=False)
    base_df['2019_Connectivity_Percentile'] = base_df['2019_Connectivity_(count_of_stations_directly_served)'].rank(ascending=False, pct=True)

    matrix_outcome_list_2 = []

    for val in base_df['Inaccessible_(1_if_not_Step_Free_Cat._A)']:

        if  val == 1:

            if val< 0.33:
                matrix_outcome_list_2.append('Top')

            elif val < 0.66:
                matrix_outcome_list_2.append('Middle')
            
            else:
                matrix_outcome_list_2.append('Bottom')
        
        else:
            matrix_outcome_list_2.append('')

    
    base_df['2019_Connectivity_Matrix_Outcome'] = matrix_outcome_list_2
    base_df['Connectivity_and_Journeys_Matrix_Outcome'] = base_df['2019_Unlocked_Journeys_Matrix_Outcome']+base_df['2019_Connectivity_Matrix_Outcome']

    #column u

    
    return base_df



    


def main():
    #any variables for the main 

    path_of_spreadsh = '/Users/kharesa-kesa.spencer/Library/CloudStorage/OneDrive-Arup/Projects/Network Rail Accessibility case/CSV WORK/Step Free Scoring_JDL_v3.00.xlsx'
    base_df  = pd.read_excel(path_of_spreadsh, sheet_name = "All Stations", header=2 , usecols="B:AS", engine='openpyxl')
    base_df.columns = [c.replace(' ','_') for c in base_df.columns]


    updated_cats_and_jrnys = get_new_categories_set_jrnys(base_df)




if __name__ == "__main__":
    main()
