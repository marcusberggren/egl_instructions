import re

import pandas as pd
import xlwings as xw


def egl_si():
    file_path = r"docs\si\RDX1 005S SEGOT.xlsm"
    df = pd.read_excel(file_path, sheet_name='Output', header=8, usecols=("A,C,E,I:K"))
    df = df.rename(columns={
        '  CONTAINER NO. '      : 'CONTAINER',
        ' PACKAGE'              : 'PKG',
        '    GWT '              : 'GWT',
        '    B/L NO.  '         : 'BOOKING',
        ' MRN No / Q-status '   : 'MRN',
        ' PARTNER'              : 'MLO'
        })
    df = df.drop(index=0)
    df = df.dropna(axis='index', how='all')
    df[['CONTAINER', 'TYPE']] = df['CONTAINER'].str.split(expand=True)
    df['BOOKING'] = df['BOOKING'].str.replace(' ', '')
    df.loc[:,'BOOKING'] = df['BOOKING'].astype(str).replace(r'\.0*', '')
    df = df.reset_index(drop=True)
    df = df.sort_values('BOOKING')
    
    return df

def clean_data(string: str) -> str:
    return re.match(r"\d*[^\.0]", string)[0]

def load_df():
    wb = xw.Book.caller()
    sheet = wb.sheets('INFO')
    data_table = sheet.range('A4').expand()
    df = sheet.range(data_table).options(pd.DataFrame, index=False, header=True).value
    df = df.loc[(df.MLO == 'EVER') & (df['LOAD STATUS'] != 'MT'), :]
    df = df[['BOOKING NUMBER', 'MLO', 'CONTAINER', 'ISO TYPE', 'NET WEIGHT', 'LOAD STATUS', 'VGM', 'OOG', 'REMARK', 'IMDG', 'UNNR', 'MRN', 'TEMP', 'PACKAGES']]
    df.insert(0, 'COMMENT', 'Bokningsblad')
    df.loc[:,'BOOKING NUMBER'] = df['BOOKING NUMBER'].astype(str).replace('.0', '', regex=False)
    df['MLO'] = None
    df['CONTAINER'] = None
    df['ISO TYPE'] = None
    df['NET WEIGHT'] = None
    df = df.sort_values('BOOKING NUMBER')
    
    return df

def export_df(df):
    file = r"C:\Users\SWV224\Documents\python_templates\template-egl-check.xlsb"

    xw.Book(file).sheets('EGL').range('C6').options(
        pd.DataFrame,
        index=False,
        header=False
        ).value = df.copy()

def fill_in_si_data(dfb: pd.DataFrame, dfe: pd.DataFrame) -> pd.DataFrame:

    def booking_count(df1: pd.DataFrame, df2: pd.DataFrame, booking: str) -> int:
        df1_count, df2_count = 0, 0

        if booking in df1['BOOKING'].values:
            df1_count = df1['BOOKING'].value_counts()[booking]
        if booking in df2['BOOKING'].values:
            df2_count = df2['BOOKING'].value_counts()[booking]
        
        return max(df1_count, df2_count)
    
    dfb = dfb.rename(columns={'BOOKING NUMBER': 'BOOKING'})

    df_concat = pd.concat([dfb, dfe])
    booking_array = df_concat['BOOKING'].unique()

    string = str()
    string += ''.join(f'{booking}, ' * booking_count(dfb, dfe, booking) for booking in booking_array)


    

    """
    df['COMMENT'] = 'OK'
    booking_cont_df = df.groupby('BOOKING NUMBER').size()
    booking_cont_df2 = df2.groupby('BOOKING')['CONTAINER'].apply(list).to_dict()
    

    for num, value in enumerate(df2.BOOKING):

        if len(booking_cont_df2[value]) > booking_cont_df[value]:
            df.loc[num + 0.5,]

        if value == df['BOOKING NUMBER'][num]:

            df_grouped[value] -= 1
            df2_grouped[value] -= 1

            df.loc[num, 'COMMENT'] = "OK"
            df.loc[num, 'MLO'] = df2.MLO[num]
            df.loc[num,'CONTAINER'] = df2.loc[num, 'CONTAINER']
            df.loc[num,'ISO TYPE'] = df2.loc[num, 'TYPE']
            df.loc[num,'NET WEIGHT'] = df2.loc[num, 'GWT']
            df.loc[num,'MRN'] = df2.loc[num, 'MRN']
            df.loc[num,'PACKAGES'] = df2.loc[num, 'PKG']
        
        elif value not in df_grouped['BOOKING NUMBER']


        elif df['BOOKING NUMBER'][num-1] == df['BOOKING NUMBER'][num] and df2.BOOKING[num-1] != df2.BOOKING[num]:
            df.loc[num, 'COMMENT'] = "Inte i EGL S/I"
        else:
            df.loc[num + 0.5,'MLO'] = df2.loc[num, 'MLO']
            df.loc[num + 0.5,'MLO'] = df2.loc[num, 'MLO']
            df.loc[num + 0.5,'CONTAINER'] = df2.loc[num, 'CONTAINER']
            df.loc[num + 0.5,'ISO TYPE'] = df2.loc[num, 'TYPE']
            df.loc[num + 0.5,'NET WEIGHT'] = df2.loc[num, 'GWT']
            df.loc[num + 0.5,'MRN'] = df2.loc[num, 'MRN']
            df.loc[num + 0.5,'PACKAGES'] = df2.loc[num, 'PKG']
            
            df = df.sort_index().reset_index(drop=True)

    return df
    """

def get_data_from_template():
    file_path = r"C:\Users\SWV224\Documents\python_templates\template-egl-check.xlsb"

    df = pd.read_excel(file_path, sheet_name='EGL', header=4)

    return df
    

if __name__ == '__main__':
    xw.Book(r'docs\bokningsblad\0114_Bokningsblad_ANINA_23004_SEGOT.xlsb').set_mock_caller()
    #df = load_df()
    #df = get_data_from_template()
    #df2 = egl_si()
    #fill_in_si_data(df, df2).to_excel('test2.xlsx')
    
    df = load_df()
    df2 = egl_si()
    """
    df_grouped = df.groupby('BOOKING NUMBER').size().reset_index(name='COUNTS')
    df2_grouped = df2.groupby('BOOKING').size().reset_index(name='COUNTS')
    """
    #print(df_grouped)
    #print(df2_grouped)

    fill_in_si_data(df, df2)


