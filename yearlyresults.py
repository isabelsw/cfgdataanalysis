#Import panda library as pd
import pandas as pd

#Get the existing Excel file
df = pd.read_excel("Book1.xlsx", sheet_name="Results Sheet")

#Add the new data and create a dataframe for it
data = {"Sales Total": [45542],
        "Sales Average": [3796],
        "Sales Maximum": [7479],
        "Sales Minimum": [1521],
        "Month with Highest Sales": ["Jul"],
        "Month with Lowest Sales": ["Feb"]
        }

df_new = pd.DataFrame(data)

#Write the dataframe to a new sheet
with pd.ExcelWriter("Book1.xlsx", engine="openpyxl", mode="a") as writer:
    df_new.to_excel(writer, sheet_name="Yearly Results", index=False)