def main(temp):
    import openpyxl

    print("temp: " + temp)

    wb = openpyxl.Workbook()


    wb.save("Our first workbook.xlsx")  # create xlsx file

    wb_1 = openpyxl.load_workbook("Our first workbook.xlsx")

    for sheet in wb_1:
        print(sheet.title)

if __name__ == "__main__":
    main()

