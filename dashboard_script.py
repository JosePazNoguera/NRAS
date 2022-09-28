'''
Authors Jose and Kharesa

Compressed script for PowerBI to directly get the df dataframe into the ODMatrix table
15/09/2022

'''



from operator import index
import pandas as pd, numpy as np, glob, ast, openpyxl, shutil, pyodbc, random, datetime
from openpyxl import Workbook
from openpyxl import load_workbook
from datetime import datetime


#random scenario number for naming convention, to be replaced by input script number
sn = str(random.randint(100, 999))

#reading the data in from access table
# connect to the access database
conn = pyodbc.connect(
    r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\jose.delapaznoguera\OneDrive - Arup\NRAS Secondment\Automation\Inputs\MOIRAOD (1).accdb;')
'''
cursor = conn.cursor()
query = cursor.execute('select * from ODMatrix')
for row in cursor.fetchall():
    print (row)
'''
query = 'select * from ODMatrixAfA'
df = pd.read_sql(query, conn)


# Add numerical score based on a dictionary
my_dict = {'0': 0, 'A': 1, 'B': 2, 'B1': 3, 'B2': 4, 'B3': 5, 'C': 6, 'Null': -1}
origin_score = df.AfAOrigin.map(my_dict)
destination_score = df.AfADest.map(my_dict)
df['origin_score'] = origin_score
df['destination_score'] = destination_score

# Remove OD pairs where the score is 0 or Null. Save the number of removed records for traceability
dropped_origins = len(df[(df.origin_score < 1) | (df.origin_score.isna())])
dropped_destinations = len(df[(df.destination_score < 1) | (df.destination_score.isna())])
df = df.drop(df[(df.origin_score < 1) | (df.origin_score.isna())].index)
df = df.drop(df[(df.destination_score < 1) | (df.destination_score.isna())].index)
# df.info()


# Categorize the journeys by looking at the maximum numerical score of the origin and destination stations
jny_score = np.maximum(df.origin_score, df.destination_score)

# Add the list as a new column to the existing dataframe
df['jny_score'] = jny_score
my_dict_2 = {1: 'A', 2: 'B', 3: 'B1', 4: 'B2', 5: 'B3', 6: 'C'}
jny_category = df.jny_score.map(my_dict_2)
df['jny_category'] = jny_category
df['Total_Journeys'] = df['STDJOURNEYS'] + df['1stJOURNEYS']

#df.info()
v1 = df.groupby("jny_category").Total_Journeys.sum()





### STATION UPGRADE ROUTINE


# Import stations to be upgraded
input_path = r"C:\Users\jose.delapaznoguera\OneDrive - Arup\NRAS Secondment\Automation\Inputs\Input template.csv"
upgrade_list = pd.read_csv(input_path)

# Transform the list into a dataframe
upgrade_list.columns = [c.replace(' ','_') for c in upgrade_list.columns]

# Data columns (total 3 columns):
#  #   Column        Non-Null Count  Dtype
# ---  ------        --------------  -----
#  0   Station       2 non-null      object
#  1   TLC           5 non-null      object
#  2   New_Category  6 non-null      object
# dtypes: object(3)



# Next steps:
# 1. find the stations to be upgraded.
for tlc in upgrade_list.TLC:
    if str(tlc) == 'nan':
        continue
    # Update category. It is necessary to use the .item() method to the Series ,
    new_category = upgrade_list.loc[upgrade_list.TLC == tlc, 'New_Category'].item()

    # Update origin category
    df.loc[df.OriginTLC == str(tlc), 'AfAOrigin'] = new_category
    # Update destination category
    df.loc[df.DestinationTLC == str(tlc), 'AfADest'] = new_category


# Update origin score
df.origin_score = df.AfAOrigin.map(my_dict)
# Update destination score
df.destination_score = df.AfADest.map(my_dict)
# Update journey score
df.jny_score = np.maximum(df.origin_score, df.destination_score)
# Update journey category
df.jny_category = df.jny_score.map(my_dict_2)
# Concat the 2 categories together
df['concat_categories'] = df.AfAOrigin + df.AfADest

#dropping for the powerbi
df.drop(axis=1,columns=['origin_score', 'destination_score', 'jny_score', 'jny_category'], inplace=True)

New_ODMatrix = df
