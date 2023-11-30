import pandas as pd
from datetime import datetime

date_pattern_long = "%Y-%m-%d %H:%M:%S"
date_pattern_short = "%m/%d/%Y"

input = 'mje_DOE_Electric_Disturbance_Events.xlsx';
df = pd.ExcelFile(input)
for sheet in df.sheet_names:
    dfs = pd.read_excel(input, sheet_name=sheet)
    #print(sheet)
    #print(dfs.columns)
    if('Area Affected' and 'Number of Customers Affected' in dfs.columns):
        for index, row in dfs.iterrows():
            if not row.isnull().values.any():
                date_field = str(row['Date'])
                if '/' in date_field:
                    parsed_date = datetime.strptime(date_field, date_pattern_short)
                elif ':' in date_field:
                    parsed_date = datetime.strptime(date_field, date_pattern_long)
                print(parsed_date, row['Date'], row['Area Affected'], row['Number of Customers Affected'])
