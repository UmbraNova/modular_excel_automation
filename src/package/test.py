import re
# testing options to filter cell values

cell_value = "  DENUMIRE PRODUS "


def test(cell_value):
    test_dict = {
        "test_key": ["denumire produs"]
    }

    for values in test_dict.values():
        for i_val in values:
            cell_value = cell_value.replace(",", " ").replace(".", " ").replace("/", " ")
            if cell_value in i_val or i_val in cell_value:
                print("cheeeeeeeeee__k")





test(str(cell_value).lower())