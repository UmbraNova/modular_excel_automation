import pandas as pd
from datetime import datetime


# item status, citeste doua coloane, una cu ITM00100 si una cu data trecuta,
# sterge data care depaseste ziua de azi, sterge duplicatele la ITM
# pastrand ultima data actuala care nu depaseste data curenta 


file_path = 'X:/AUR/11.2024/25.11.2024/item status - grup/Default28_11_2024_10_16_50.xlsx'
output_path = 'X:/AUR/11.2024/25.11.2024/item status - grup/updated_item_status.xlsx'

df = pd.read_excel(file_path)
df.columns = df.columns.str.strip()
column_a = 'A'
column_f = 'F'
index_a = ord(column_a) - ord('A')
index_f = ord(column_f) - ord('A')

df.iloc[:, index_f] = pd.to_datetime(df.iloc[:, index_f], errors='coerce')  # Convert column F to datetime
current_date = pd.to_datetime(datetime.now().date())
df = df[df.iloc[:, index_f] <= current_date]  # Delete rows where the date in column F is later than the current date
df = df.loc[df.groupby(df.columns[index_a])[df.columns[index_f]].idxmax()]  # Remove duplicates based on column A, keeping the row with the most recent date in column F
df.iloc[:, index_f] = df.iloc[:, index_f].astype(str).str[:10]  # Convert the date in column F to string format and slice to keep only the date

df.to_excel(output_path, index=False)

print("Cleaning complete. The cleaned file is saved as:", output_path)
