import pandas as pd

from convert import generate_docx


def extract_data_from_excell(file_name: str, sheet_name: str) -> None:
    result = []
    df = pd.read_excel(file_name, sheet_name=sheet_name)
    for main_row_index in range(3, len(df), 9):
        main_index, main_row = tuple(df.iterrows())[main_row_index]
        pds_no = tuple(main_row.items())[1][1]
        product_name = tuple(main_row.items())[2][1]
        product_description = tuple(main_row.items())[3][1]
        benefits = tuple(main_row.items())[4][1]
        benefits = benefits.split('\n')
        sae_grade = tuple(main_row.items())[6][1]
        api = tuple(main_row.items())[7][1]
        ilsac = tuple(main_row.items())[8][1]
        acea = tuple(main_row.items())[9][1]
        oem_claims = tuple(main_row.items())[10][1]
        recommendations = tuple(main_row.items())[11][1]

        performance_claims = get_performance_claims(api, ilsac, acea, oem_claims, recommendations)

        row_item = {
            'pds_no': pds_no,
            'product_name': product_name,
            'product_description': product_description,
            'benefits': benefits,
            'sae_grade': sae_grade,
            'performance_claims': performance_claims,
        }
        result.append(row_item)
    return result


def get_performance_claims(
        api: str,
        ilsac: str,
        acea: str,
        oem_claims: str,
        recommendations: str
        ) -> list:
    result = []
    if api not in ['Not available', 'Not applicable', '', None]:
        result.append(f'API {api}')
    if ilsac not in ['Not available', 'Not applicable', '', None]:
        result.append(f'ILSAC {ilsac}')
    if acea not in ['Not available', 'Not applicable', '', None]:
        result.append(f'ACEA {acea}')
    if oem_claims not in ['Not available', 'Not applicable', '', None]:
        if ',' in oem_claims:
            oem_claims = oem_claims.split(',')
        else:
            oem_claims = oem_claims.split('\n')
        result.extend([item.strip() for item in oem_claims])
    if recommendations not in ['Not available', 'Not applicable', '', None]:
        if ',' in recommendations:
            recommendations = recommendations.split(',')
        else:
            recommendations = recommendations.split('\n')
        result.extend([item.strip() for item in recommendations])
    return result


excell_files = extract_data_from_excell('excell_tom.xlsx', 'Engine oil')


for _ in excell_files:
    generate_docx(
        'tom_test.docx',
        _['pds_no'],
        _['product_name'],
        _['product_description'],
        _['benefits'],
        _['sae_grade'],
        'Engine oil',
        _['performance_claims'],
        )
