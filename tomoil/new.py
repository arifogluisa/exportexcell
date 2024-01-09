import pandas as pd


def extract_data_from_excell(file_name: str, sheet_name: str) -> None:
    result = []
    df = pd.read_excel(file_name, sheet_name=sheet_name)
    for main_row_index in range(len(df)):
        main_index, main_row = tuple(df.iterrows())[main_row_index]
        product_code = tuple(main_row.items())[0][1]
        product_name = tuple(main_row.items())[1][1]
        category = tuple(main_row.items())[2][1]

        row_item = {
            'product_code': product_code,
            'product_name': product_name,
            'category': category,
        }
        result.append(row_item)
    return result


print(extract_data_from_excell('Aminol.xlsx', 'Sheet1'))
