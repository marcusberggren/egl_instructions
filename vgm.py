import os
import tkinter as tk
from pathlib import Path
from tkinter import filedialog

import pandas as pd
import xlwings as xw


def vgm_instructions(vgm_file):
    home = str(Path.home())
    dir = f'BOLLORE\XPF - Documents\SERVICES\RGX'
    init_dir = os.path.join(home, dir)

    if not vgm_file:
        vgm_file = open_vgm(init_dir)

    df1 = load_caller_df()
    df2 = egl_vgm()

    df_concat = pd.concat([df1, df2])
    booking_array = df_concat['BOOKING NUMBER'].unique()

    array_of_strings = ""
    array_of_strings += ''.join(f'{booking},' * booking_count(df1, df2, booking) for booking in booking_array)

    booking_list = ""
    booking_list = array_of_strings.split(',')
    booking_list.pop()

    df4 = pd.DataFrame(columns=df1.columns)
    df4['BOOKING NUMBER'] = booking_list

    counter_df1 = 0
    counter_df2 = 0

    for num, booking in enumerate(df4['BOOKING NUMBER']):

        if not counter_df1 == len(df1):
            if booking == df1['BOOKING NUMBER'][counter_df1]:
                df4['COMMENT'][num] = df1['COMMENT'][counter_df1]
                df4['LOAD STATUS'][num] = df1['LOAD STATUS'][counter_df1]
                counter_df1 += 1

        if not counter_df2 == len(df2):
            if booking == df2['BOOKING NUMBER'][counter_df2]:
                df4['CONTAINER'][num] = df2['CONTAINER'][counter_df2]
                df4['PACKAGES'][num] = df2['PACKAGES'][counter_df2]
                df4['NET WEIGHT'][num] = df2['NET WEIGHT'][counter_df2]
                df4['MRN'][num] = df2['MRN'][counter_df2]
                df4['MLO'][num] = df2['MLO'][counter_df2]
                counter_df2 += 1

    #xw.Book.caller().sheets('EGL').range('B6').value = df4.copy()

    df4.to_excel('test_vgm.xlsx')


def egl_vgm(file_path):
    df = pd.read_excel(file_path, sheet_name='Raw Data', header=8, usecols=("A,B,D"))
    df = df.rename(columns={
        'Booking No.'   : 'BOOKING NUMBER',
        'Container No.' : 'CONTAINER',
        'VGM'           : 'VGM'
        })
    df = df.sort_values('BOOKING NUMBER').reset_index(drop=True)
    return df

def open_vgm(path):
    root = tk.Tk()
    root.lift()
    root.withdraw()

    file = filedialog.askopenfile(initialdir=path, title="select VGM-file", filetypes=[("Excel files", ".xls")])
    root.quit()

    if not file:
        exit()

    return file


def load_caller_df():
    wb = xw.Book.caller()
    sheet = wb.sheets('EGL')
    data_table = sheet.range('B5').expand()
    df = sheet.range(data_table).options(pd.DataFrame, index=False, header=True).value
    df = df.sort_values('BOOKING NUMBER').reset_index(drop=True)

    return df

def booking_count(df1: pd.DataFrame, df2: pd.DataFrame, booking: str) -> int:
        df1_count, df2_count = 0, 0

        if booking in df1['BOOKING NUMBER'].values:
            df1_count = df1['BOOKING NUMBER'].value_counts()[booking]
        if booking in df2['BOOKING NUMBER'].values:
            df2_count = df2['BOOKING NUMBER'].value_counts()[booking]

        return max(df1_count, df2_count)


if __name__ == '__main__':
    file_path = r"docs\si\RDX1 005S SEGOT.xlsm"
    vgm_instructions(file_path)