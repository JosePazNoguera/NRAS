"""

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

import pandas as pd, numpy as np, glob, ast

# Change the settings to output thousand separators: Use f'{value:n}' For Python ≥3.6

# Set the path where the data is saved. The code will pick all csv files
path = r'C:/Users/jose.delapaznoguera/Projects/NRAS/Data'
filenames = glob.glob(path + "/*.csv")

# Create an temporary list to store the content of each file
my_list = []
for filename in filenames:
    if filename == 'C:/Users/jose.delapaznoguera/Projects/NRAS/Data\\Input template.csv' or filename == 'C:/Users/jose.delapaznoguera/Projects/NRAS/Data\\Demographics.csv':
        continue
    my_list.append(pd.read_csv(filename))
#    print(filename)


# Transform the list into a dataframe
my_df = pd.concat(my_list, ignore_index=True)
my_df.columns = [c.replace(' ', '_') for c in my_df.columns]

# list of the columns with nan values
nan_cols = my_df.loc[:, my_df.isna().any(axis=0)]

# if there is only one column check it isnt region, region can be ignored - metadata
if len(nan_cols.columns) == 0:
    if nan_cols.columns[0] == 'Region':
        print('only region has empty rows, proceeding')

# else drop all rows within the nan columns with empty rows
else:
    # if 'Total_Journeys' in nan_cols.columns:

    # first does all rows with no data
    my_df.dropna(axis=0, how='all', subset=nan_cols.columns)

    # then does all rows with no data
    my_df.dropna(axis=0, how='any', subset=nan_cols.columns)

# Make sure we created a dataframe
# print(type(my_df))
# my_df.info()
#
# Data columns (total 6 columns):
#  #   Column                Non-Null Count    Dtype
# ---  ------                --------------    -----
#  0   Origin_TLC            1429764 non-null  object
#  1   Destination_TLC       1429764 non-null  object
#  2   Origin_Category       1429764 non-null  object
#  3   Destination_Category  1429764 non-null  object
#  4   Total_Journeys        1429764 non-null  int64
#  5   Region                1404689 non-null  object
# dtypes: int64(1), object(5)

# Print the first few rows to take a look. Tick the option to show all the columns
pd.set_option('display.max_columns', None)
# print(my_df.head(15))

# Take a look at the current totals per category
# print(my_df.groupby("Origin_Category").Origin_Category.count())
# print(my_df.groupby("Destination_Category").Destination_Category.count())


# create a copy of our dataframe
base_df = my_df.copy()
# Add numerical score based on a dictionary
my_dict = {'0': 0, 'A': 1, 'B': 2, 'B1': 3, 'B2': 4, 'B3': 5, 'C': 6, 'Null': -1}
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

print(
    f"A total of {dropped_origins:n} origins and {dropped_destinations:n} destinations were invalid and not included in the analysis.")

# Categorize the journeys by looking at the maximum numerical score of the origin and destination stations
jny_score = np.maximum(base_df.origin_score, base_df.destination_score)

# Add the list as a new column to the existing dataframe
base_df['jny_score'] = jny_score
my_dict_2 = {1: 'A', 2: 'B', 3: 'B1', 4: 'B2', 5: 'B3', 6: 'C'}
jny_category = base_df.jny_score.map(my_dict_2)
base_df['jny_category'] = jny_category
# print(base_df.tail(10))

# base_df.info()
v1 = base_df.groupby("jny_category").Total_Journeys.sum()
# print(v1)
# print(type(v1))
# print(base_df.tail(10))

### STATION UPGRADE ROUTINE

# Import stations to be upgraded
input_path = r'C:/Users/jose.delapaznoguera/Projects/NRAS/Data/Input template.csv'
upgrade_list = pd.read_csv(input_path)

# Transform the list into a dataframe
upgrade_list.columns = [c.replace(' ', '_') for c in upgrade_list.columns]
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

# Concat the 2 categories together
concat_categories = scenario_1.Origin_Category + scenario_1.Destination_Category
scenario_1['concat_categories'] = concat_categories

# We need to select the journeys where one end is accessible and the other is not.
# Step 1: select jnys where at least 1 end is accessible
scenario_1_clean = scenario_1.loc[(scenario_1.Origin_Category == 'A') |
                                  (scenario_1.Origin_Category == 'B1') |
                                  (scenario_1.Destination_Category == 'A') |
                                  (scenario_1.Destination_Category == 'B1')
]

# Step 2: remove jnys where both ends are accessible
scenario_1_clean = scenario_1_clean.loc[(scenario_1_clean.concat_categories != 'AA') &
                                        (scenario_1_clean.concat_categories != 'AB1') &
                                        (scenario_1_clean.concat_categories != 'B1A') &
                                        (scenario_1_clean.concat_categories != 'B1B1')
                                        ]

#
# # print(base_df.info())
# print(scenario_1.info())

# grouping by TLC and cat and totalling journeys, setting all Cat A as None
grouped_origin_df = (scenario_1_clean.groupby(["Origin_TLC", "Origin_Category"])["Total_Journeys"].sum()).to_frame()
grouped_origin_df.reset_index(inplace=True)
grouped_origin_df.loc[grouped_origin_df.Origin_Category == 'A', 'Total_Journeys'] = None
grouped_origin_df.loc[grouped_origin_df.Origin_Category == 'B1', 'Total_Journeys'] = None

grouped_destination_df = (
    scenario_1_clean.groupby(["Destination_TLC", "Destination_Category"])["Total_Journeys"].sum()).to_frame()
grouped_destination_df.reset_index(inplace=True)
grouped_destination_df.loc[grouped_destination_df.Destination_Category == 'A', 'Total_Journeys'] = None
grouped_destination_df.loc[grouped_destination_df.Destination_Category == 'B1', 'Total_Journeys'] = None

# saving these to csv
grouped_origin_df.to_csv('Outputs/origin_grouped_By.csv')
grouped_destination_df.to_csv('Outputs/destination_grouped_By.csv')
scenario_1_clean.to_csv('Outputs/scen_1_clean.csv')
