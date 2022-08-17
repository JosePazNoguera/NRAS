import pandas as pd
from IPython.display import display



# initialise data of lists.
colors = {'first_set': ['99', '88', '77', '66',
						'55', '44', '33', '22'],
		'second_set': ['1', '2', '3', '4', '5',
						'6', '7', '8']
		}

color = {'first_set': ['a', 'b', 'c', 'd', 'e',
						'f', 'g', 'h'],
		'second_set': ['VI', 'IN', 'BL', 'GR',
						'YE', 'OR', 'RE', 'WI']
		}
# Calling DataFrame constructor on list
df = pd.DataFrame(colors, columns=['first_set', 'second_set'])
df1 = pd.DataFrame(color, columns=['first_set', 'second_set'])

# Display the Output
display(df)
display(df1)

# selecting old value
a = df1['first_set'][4]

# selecting new value
b = df['first_set'][1]


# replace values of one DataFrame
# with the value of another DataFrame
df1 = df1.replace(a,b)

# Display the Output
display(df1)
