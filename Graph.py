#Import: seaborn as sns; pandas as pd; matplotlib as plt
import seaborn as sns
import pandas as pd
import matplotlib.pyplot as plt

#Specify columns from which to retrieve the data
require_cols = [2, 5]

#Read only the specific columns from the excel file
required_df = pd.read_excel('Book1.xlsx', sheet_name="Results Sheet", usecols=require_cols, index_col=0)

#Print the data
print(required_df)

#Create a simple lineplot using seaborn
sns.lineplot(required_df)

#Show this lineplot using matplotlib
plt.show()