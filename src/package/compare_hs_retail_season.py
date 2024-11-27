import openpyxl


wb = openpyxl.load_workbook("X:/AUR/11.2024/25.11.2024/corelatie,hs,retail,season/codes in system.xlsx")
ws = wb.active



count = 0  # for testing

resulted_core = {}

for row in ws.iter_rows():
    cell_value = ""
    first_col = None
    for cell in row[2:5]:
        cell_value = str(cell.value) if cell.value else ""
        if cell_value.isdigit():
            if len(cell_value) > 7:
                first_col = cell_value
                if cell_value not in resulted_core.keys():
                    resulted_core[first_col] = list()
            elif len(cell_value) > 6:  # hs codes
                if first_col is not None:
                    if cell_value not in resulted_core[first_col]:
                        resulted_core[first_col].append(cell_value)
            elif 2 < len(cell_value) <= 6:  # retail codes
                if first_col is not None:
                    if cell_value not in resulted_core[first_col]:
                        resulted_core[first_col].append(cell_value)


    # if count == 7:  # limit nr of rows for testing
    #     break
    # count+=1


print("-"*30)
for keys, values in resulted_core.items():
    if values:
        print(keys, values)

















