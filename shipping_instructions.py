import pandas as pd
import xlwings as xw

def egl_si():
    file_path = r"docs\si\RDX1 005S SEGOT.xlsm"
    df = pd.read_excel(file_path, sheet_name='Output', header=8, usecols=("A,C,E,I:K"))
    df = df.rename(columns={
        '  CONTAINER NO. '      : 'CONTAINER_NO',
        ' PACKAGE'              : 'PKG',
        '    GWT '              : 'GWT',
        '    B/L NO.  '         : 'BOOKING_NO',
        ' MRN No / Q-status '   : 'MRN',
        ' PARTNER'              : 'MLO'
        })
    df = df.drop(index=0)
    df = df.dropna(axis='index', how='all')
    df[['CONTAINER_NO', 'TYPE']] = df.CONTAINER_NO.str.split(expand=True)
    df.BOOKING_NO = df.BOOKING_NO.str.replace(' ', '')
    df.BOOKING_NO = df.BOOKING_NO.astype(str).astype(float)
    df.reset_index(drop=True)
    first_col = df.pop('BOOKING_NO')
    df.insert(0, 'BOOKING_NO', first_col)
    
    return pd.DataFrame(df)

def egl_vgm():
    file_path = r"docs\vgm\SB1P9XFQ1.XLS"

    df = pd.read_excel(file_path, 'Raw Data', header=8, usecols=("A,B,D"))
    df = df.rename(columns={
        'Booking No.'    : 'BOOKING_NO',
        'Container No.'  : "CONTAINER_NO"
        })

    return df

def load_df():
    wb = xw.Book.caller()
    sheet = wb.sheets('INFO')
    data_table = sheet.range('A4').expand()
    df = sheet.range(data_table).options(pd.DataFrame, index=False, header=True).value
    df = df.loc[(df.MLO == 'EVER') & (df['LOAD STATUS'] != 'MT'), :]
    df = df.rename(columns={'BOOKING NUMBER': 'BOOKING_NO'})
    df = df[['BOOKING_NO', 'REMARK', 'IMDG', 'UNNR', 'TEMP']]
    
    #df['BOOKING_NO'] = df['BOOKING_NO'].values.astype(str)
    df['BOOKING_NO'] = df['BOOKING_NO'].apply(lambda f: format(f, '.0f'))
    df['BOOKING_NO'] = df['BOOKING_NO'].astype(str).astype(float)

    return df

def export_df(df):
    file = r"C:\Users\SWV224\Documents\python_templates\template-egl-check.xlsb"

    xw.Book(file).sheets('EGL').range('B6').options(
        pd.DataFrame,
        index=False,
        header=False
        ).value = df.copy()

    

if __name__ == '__main__':
    xw.Book(r'docs\bokningsblad\0114_Bokningsblad_ANINA_23004_SEGOT.xlsb').set_mock_caller()




