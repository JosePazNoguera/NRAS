import pandas as pd, numpy as np, glob

df = pd.DataFrame([[1, 2], [4, 5], [7, 8]], index=['cobra', 'viper', 'sidewinder'], columns=['max_speed', 'shield'])
print(df)

jose = "A"
df.loc[df.max_speed == 7, 'shield'] = jose

print(df)