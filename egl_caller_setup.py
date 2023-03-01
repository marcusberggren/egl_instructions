import os
from pathlib import Path

import pandas as pd
import xlwings as xw


def main():
    
    df = load_caller_df()
    with xw.App() as app:
        wb = app.books.open(template_file())
        wb.sheets('EGL').range('B6').options(pd.DataFrame, index=False, header=False).value = df.copy()
        wb.save()


def template_file():
    home = str(Path.home())
    template = r'Documents\python_templates\template-egl-instructions.xlsb'
    template_path = os.path.join(home, template)
    return template_path


def load_caller_df():
    wb = xw.Book.caller()
    sheet = wb.sheets('INFO')
    data_table = sheet.range('A4').expand()
    df = sheet.range(data_table).options(pd.DataFrame, index=False, header=True).value
    df = df.loc[(df['MLO'] == 'EVER') & (df['LOAD STATUS'] != 'MT'), :]
    df = df[['BOOKING NUMBER', 'MLO', 'CONTAINER', 'ISO TYPE', 'NET WEIGHT', 'LOAD STATUS', 'VGM', 'OOG', 'REMARK', 'IMDG', 'UNNR', 'MRN', 'TEMP', 'PACKAGES']]
    df = pd.DataFrame(df)
    df['BOOKING NUMBER'] = df['BOOKING NUMBER'].astype(str)
    df['BOOKING NUMBER'] = df['BOOKING NUMBER'].str.replace('.0', '', regex=False)
    df.insert(0, 'COMMENT', 'Bokningsblad')
    df.loc[:,'BOOKING NUMBER'] = df['BOOKING NUMBER'].astype(str).replace('.0', '', regex=False)
    df['MLO'] = None
    df['CONTAINER'] = None
    df['ISO TYPE'] = None
    df['NET WEIGHT'] = None
    df['COMMENT'] = "Bokningsblad"
    df = df.sort_values('BOOKING NUMBER').reset_index(drop=True)

    return df

if __name__ == '__main__':
    xw.Book(r'docs\0114_Bokningsblad_ANINA_23004_SEGOT-ursprunglig.xlsb').set_mock_caller()
    main()