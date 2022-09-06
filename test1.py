import pandas as pd

# load dataset into Pandas DataFrame
demog = pd.read_csv("Data\Demographics.csv",header=0, index_col= 0)
# print(demog.head(10))
# print(demog.columns)

y = demog.iloc[:, -1:]
demog.drop(['Employment Deprivation Rate (% of working population)'],axis = 1, inplace=True)
# print(y)
# print(demog.columns)

# Categorize the non numeric columns
# demog["Region"] = demog["Region"].astype('category')
# demog["Region_cat"] = demog["Region"].cat.codes

# Drop the non numeric columns from the dataframe
demog.drop(['Station Name', 'Region', 'Country', 'Christian (%)',
            'Buddhist (%)', 'Hindu (%)', 'Jewish (%)', 'Muslim (%)', 'Sikh (%)',
            'Other Religion (%)', 'No religion (%)', 'Religion not stated (%)'], axis=1, inplace=True)

df = y.join(demog)
print(df.columns)
df.info()

import numpy as np
from matplotlib import pyplot as plt
from sklearn.datasets.samples_generator import make_blobs
from sklearn.cluster import KMeans

