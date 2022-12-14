"""
NRAS
Authors: Jose De la Paz Noguera and Kharesa-Kesa Spencer
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

from operator import index
import pandas as pd, numpy as np, glob, ast, openpyxl, shutil, pyodbc, random, datetime
from openpyxl import Workbook
from openpyxl import load_workbook
from datetime import datetime


def reading_input():
    
    #temporary variable to direct the code, when the access database is updated then we can remove

    #connect to the access database
    conn = pyodbc.connect('Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=;C:/Users/jose.delapaznoguera/OneDrive - Arup/NRAS Secondment/Automation/Inputs/MOIRAOD (1).accdb')
    '''
    cursor = conn.cursor()
    query = cursor.execute('select * from ODMatrix')
    for row in cursor.fetchall():
        print (row)
    '''
    query = 'select * from ODMatrix'
    my_df = pd.read_sql(query, conn)

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

def calculations(base_df, sn, target):
    
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

    #reading from the spreadsheet
    st_cat_df = pd.read_excel(target, sheet_name = "St_Cat", engine='openpyxl')

    st_cat_df = st_cat_df.loc[:,~st_cat_df.columns.duplicated()].copy()


    #st_cat_df = df[['CRS Code','Station Name (MOIRA Name)','Category','Region']]
    st_cat_df.rename(columns={'CRS Code': 'CRS_Code', 'Station Name (MOIRA Name)': 'Station_Name'},inplace=True)
    #dropping duplicate of crs code


    # Import stations to be upgraded
    input_path = 'C:/Users/jose.delapaznoguera/OneDrive - Arup/NRAS Secondment/Automation/Inputs/Input template.csv'
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

        #if the current tlc is in the station category spreadsheet then upgrade st_cat_df with new category
        if tlc in st_cat_df.values:
            new_category1 = upgrade_list.loc[upgrade_list.TLC == tlc, 'New_Category'].item()
            st_cat_df.loc[st_cat_df.CRS_Code == str(tlc), 'Category'] = new_category1

    #export back to csv
    with pd.ExcelWriter(target, mode="a",engine="openpyxl",if_sheet_exists="replace",) as writer:
        st_cat_df.to_excel(writer, sheet_name="St_Cat", index=False)


    # Update origin score
    scenario_1.origin_score = scenario_1.Origin_Category.map(my_dict)
    # Update destination score
    scenario_1.destination_score = scenario_1.Destination_Category.map(my_dict)
    # Update journey score
    scenario_1.jny_score = np.maximum(scenario_1.origin_score, scenario_1.destination_score)
    # Update journey category
    scenario_1.jny_category = scenario_1.jny_score.map(my_dict_2)

    # Concat the 2 categories together
    scenario_1['concat_categories'] = scenario_1.Origin_Category + scenario_1.Destination_Category

    #Outputting the log into a text file 
    output_to_log(upgrade_list, sn)

    return scenario_1

def output_to_log(upgrade_list, sn ):
    
    now = datetime.now()
    dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
    
    with open("C:/Users/jose.delapaznoguera/OneDrive - Arup/NRAS Secondment/Automation/scenarios_output_log.txt", "a+") as file_object:
        # Move read cursor to the start of file.
        file_object.seek(0)
        # If file is not empty then append '\n'
        data = file_object.read(100)
        if len(data) > 0 :
            file_object.write("\n")
        # Append text at the end of file
        file_object.write("\n")
        line = 'Scenario number: ' + sn + ' run at ' + dt_string
        file_object.write(line)
        file_object.write("\n")
        file_object.write(str(upgrade_list))
        file_object.write("\n")

def into_stepfree_spreadsheet(scenario_1, target):
    

    # We need to select the journeys where one end is accessible and the other is not.
    # Step 1: select jnys where at least 1 end is accessible
    scenario_1_clean = scenario_1.loc[(scenario_1.Origin_Category =='A') | (scenario_1.Origin_Category =='B1')
                | (scenario_1.Destination_Category =='A') | (scenario_1.Destination_Category =='B1')]

    # Step 2: remove jnys where both ends are accessible
    scenario_1_clean = scenario_1_clean.loc[(scenario_1_clean.concat_categories != 'AA') &
                                            (scenario_1_clean.concat_categories != 'AB1') &
                                            (scenario_1_clean.concat_categories != 'B1A') &
                                            (scenario_1_clean.concat_categories != 'B1B1')]


    #grouping by TLC and cat and totalling journeys, setting all Cat A as None
    grouped_origin_df= (scenario_1_clean.groupby(["Origin_TLC","Origin_Category"])["Total_Journeys"].sum()).to_frame()
    grouped_origin_df.reset_index(inplace=True)
    grouped_origin_df.loc[grouped_origin_df.Origin_Category=='A', 'Total_Journeys'] = None
    grouped_origin_df.loc[grouped_origin_df.Origin_Category=='B1', 'Total_Journeys'] = None 


    grouped_destination_df= (scenario_1_clean.groupby(["Destination_TLC","Destination_Category"])["Total_Journeys"].sum()).to_frame()
    grouped_destination_df.reset_index(inplace=True)
    grouped_destination_df.loc[grouped_destination_df.Destination_Category=='A', 'Total_Journeys'] = None 
    grouped_destination_df.loc[grouped_destination_df.Destination_Category=='B1', 'Total_Journeys'] = None 


    #Save to CSVs
    #grouped_origin_df.to_csv('origin_grouped_By.csv')
    #grouped_destination_df.to_csv('destination_grouped_By.csv')



    '''
    #the clone is then loaded into pandas directly with the sheet name defined
    workbook_origin = pd.read_excel('/Users/kharesa-kesa.spencer/Library/CloudStorage/OneDrive-Arup/Projects/Network Rail Accessibility case/CSV WORK/step_free_clone.xlsx', 
        sheet_name='Inaccessible O Accessi D')

    workbook_destination = pd.read_excel(
        '/Users/kharesa-kesa.spencer/Library/CloudStorage/OneDrive-Arup/Projects/Network Rail Accessibility case/CSV WORK/step_free_clone.xlsx', 
        sheet_name='Accessible O Inaccessi D')
    '''
    #then you can write directly to the sheet 

    '''
    # create excel writer
    writer = pd.ExcelWriter('/Users/kharesa-kesa.spencer/Library/CloudStorage/OneDrive-Arup/Projects/Network Rail Accessibility case/CSV WORK/step_free_clone.xlsx')
    # write dataframe to excel sheet named 'marks'
    grouped_origin_df.to_excel(writer, 'Accessible O Inaccessi D')
    # save the excel file
    writer.save()
   
    grouped_origin_df.to_excel('/Users/kharesa-kesa.spencer/Library/CloudStorage/OneDrive-Arup/Projects/Network Rail Accessibility case/CSV WORK/Step Free Scoring_JDL_v3.00_clone.xlsx', 
        sheet_name='Accessible O Inaccessi D', engine='openpyxl')

    grouped_destination_df.to_excel('/Users/kharesa-kesa.spencer/Library/CloudStorage/OneDrive-Arup/Projects/Network Rail Accessibility case/CSV WORK/Step Free Scoring_JDL_v3.00_clone.xlsx', 
        sheet_name='Inaccessible O Accessi D', engine='openpyxl')

    https://openpyxl.readthedocs.io/en/stable/usage.html
    '''


    #read in the workbook and then write and replace the sheets
    with pd.ExcelWriter(target, mode="a",engine="openpyxl",if_sheet_exists="replace",) as writer:
       grouped_origin_df.to_excel(writer, sheet_name="Inaccessible O Accessi D")
       grouped_destination_df.to_excel(writer, sheet_name="Accessible O Inaccessi D") 


#this is the group by function result is directly exported into the sheet


#get a list of stations missing in the dataframe but present in the spreadsheet



def main():

    #random scenario number for naming convention, to be replaced by input script number
    #sn = str(random.randint(100, 999))
    sn = 'sc_1'

    #this section here reads in the origin spreadsheet and then copies it and saves it as a clone

    original = 'C:/Users/jose.delapaznoguera/OneDrive - Arup/NRAS Secondment/Automation/Step Free Scoring_JDL_v3.00.xlsx'
    target = 'C:/Users/jose.delapaznoguera/OneDrive - Arup/NRAS Secondment/Automation/Step Free Scoring_JDL_v3.00_clone_'+sn+'.xlsx'

    #copying file files
    shutil.copyfile(original, target)


    #Initalising the base df through the inputs
    base_df = reading_input()
    print('Database loaded in successfully of shape: ', base_df.shape)


    # removing O-D pairs with 0 or null and categorising journeys 
    scenario_1 = calculations(base_df, sn, target)


    into_stepfree_spreadsheet(scenario_1, target)


    

'''
scenario_2 = pd.merge(scenario_1, test_df[['CRS_Code','Including CP6 AfA']], how='left', left_on='Origin_TLC', right_on='CRS_Code')

scenario_2['Accessible'] = scenario_2['concat_categories'].replace(['CB3','B3C','CC','B3B3'],'0')

scenario_2['Accessible'] = scenario_2['Accessible'].replace(lol,'1')

scenario_2['Accessible'] = scenario_2['Accesible'].replace(['AA','AB1','B1B1','B1A'],'2')


(scenario_2.groupby('Accessible')['Total_Journeys'].sum())/scenario_2['Total_Journeys'].sum()


not_fully_acc = ['BB3',
 'B1B2',
 'CC',
 'B3B2',
 'B1C',
 'B2B2',
 'CA',
 'B1B',
 'B2B1',
 'BB2',
 'BC',
 'B3B',
 'BB',
 'CB2',
 'B3A',
 'B2A',
 'AC',
 'B3B3',
 'AB',
 'CB3',
 'B2B3',
 'AB3',
 'AB2',
 'B2C',
 'BA',
 'BB1',
 'CB1',
 'B1B3',
 'B3B1',
 'B3C',
 'B2B',
 'CB']


 full_access = ['AA','AB1','B1B1','B1A']


'''


if __name__ == "__main__":
    main()
