import pandas as pd

file_one = 'X:/AUR/modular_excel_test/file_one.xlsx'
file_two = 'X:/AUR/modular_excel_test/file_two.xlsx'

df1 = pd.read_excel(file_one)
df2 = pd.read_excel(file_two)

column_one = 'Text_Column'
column_two = 'Search_Column'

# Ensure no NaN values in the columns
df1[column_one] = df1[column_one].fillna("")
df2[column_two] = df2[column_two].fillna("")

# Function to calculate match score for each row in file two
def calculate_score(row, words):
    score = 0
    for word in words:
        if word in row:
            score += 1
    return score

# # Process each row in file one
# for index, text in df1[column_one].items():
#     words = set(text.split())  # Split text into unique words
#     scores = df2[column_two].apply(lambda x: calculate_score(x, words))
#     df2[f"Match_Score_{index}"] = scores  # Add a score column for this row

#     # Find the row with the maximum score
#     max_row = scores.idxmax()
#     max_score = scores[max_row]
#     match_text = df2.at[max_row, column_two]

#     # Add a column with the best match from file one
#     df2[f"Best_Match_{index}"] = [match_text if i == max_row else "" for i in range(len(df2))]

# Save the updated file two
# df2.to_excel('X:/AUR/modular_excel_test/file_two_updated.xlsx', index=False)
print("Matching complete! Updated file saved as 'file_two_updated.xlsx'.")













