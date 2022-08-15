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

import pandas as pd, numpy as np, glob, math

# Change the settings to output thousand separators: Use f'{value:n}' For Python ≥3.6

# Set the path where the data is saved. The code will pick all csv files
path = r'C:/Users/jose.delapaznoguera/Projects/NRAS/Data'
filenames = glob.glob(path + "/*.csv")

# Create an temporary list to store the content of each file
my_list = []
for filename in filenames:
    if filename == 'C:/Users/jose.delapaznoguera/Projects/NRAS/Data\Input template.csv':
        continue
    my_list.append(pd.read_csv(filename))
#    print(filename)


# Transform the list into a dataframe
my_df = pd.concat(my_list, ignore_index=True)
my_df.columns = [c.replace(' ','_') for c in my_df.columns]

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

# Add numerical score based on a dictionary
base_df = my_df
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
input_path = r'C:/Users/jose.delapaznoguera/Projects/NRAS/Data\Input template.csv'
upgrade_list = pd.read_csv(input_path)

# Transform the list into a dataframe
upgrade_list.columns = [c.replace(' ','_') for c in upgrade_list.columns]
# print(upgrade_list.head())

# Create a copy of the base data frame
scenario_1 = base_df.copy()
# print(type(scenario_1))

# Next steps:
# 1. find the stations to be upgraded.
for station in upgrade_list.TLC:
    if str(station) == 'nan':
        continue
    # Update category. It is necessary to convert the series from the slicing to a string,
    new_category = upgrade_list.New_Category[upgrade_list.TLC == station]
    print(new_category)
    print(type(new_category)) # it's a series
    # This line doesn't work if I use new category
    scenario_1.loc[scenario_1.Origin_TLC == station,  'Origin_Category'] = "ABC"

# for station in upgrade_list.TLC:
#
#     # we may add a line to handle when there is no TLC but a station name is provided
#     if str(station) == 'nan':
#         continue
#     # Update category. It is necessary to convert the series from the slicing to a string
#     new_category = upgrade_list[upgrade_list.TLC == station].New_Category
#     print(new_category)
#     print(type(new_category))
#     scenario_1.loc[scenario_1.Origin_TLC == station,  'Origin_Category'] = "A"
#     # scenario_1.loc[scenario_1.Origin_TLC == station, 'origin_score'] = scenario_1.Origin_Category.map(my_dict)
    # scenario_1['Destination_TLC' == station].Destination_Category = upgrade_list['TLC' == station].New_Category
    # scenario_1['Destination_TLC' == station].destination_score = scenario_1.Destination_Category.map(my_dict)
    # scenario_1['Origin_TLC' == station].jny_score = np.maximum(scenario_1.origin_score, scenario_1.destination_score)
    # scenario_1['Destination_TLC' == station].jny_score = np.maximum(scenario_1.origin_score, scenario_1.destination_score)

# print(base_df.info())
# print(scenario_1.info())
print(base_df.loc[base_df.Origin_TLC == "HUR",  ['Origin_TLC', 'Origin_Category', 'origin_score', 'jny_category']])

print(scenario_1.loc[scenario_1.Origin_TLC == "HUR",  ['Origin_TLC', 'Origin_Category', 'origin_score', 'jny_category']])

# 2. upgrade the category, the station score and the journey score
# 3. generate the new OD matrix. It must have the same format as the original OD matrix
# 4. calculate the network scores before and after