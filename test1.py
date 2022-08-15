import pandas as pd, numpy as np, glob

# Import stations to be upgraded
input_path = r'C:/Users/jose.delapaznoguera/Projects/NRAS/Data\Input template.csv'
upgrade_df = pd.read_csv(input_path)

# Transform the list into a dataframe
upgrade_df.columns = [c.replace(' ','_') for c in upgrade_list.columns]



