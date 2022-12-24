# from openpyxl import load_workbook
import pandas as pd
import lasio, math


excel_wells_file = 'wells_data.xlsx'
sheet_name = 'Лист 1'
CURVES_MNEMONIC = ('GK', 'IK', 'NKT', 'NKTB')   # Варианты возможных мнемоник кривых которые понадобятся для расчетов
ln = math.log1p                                 # Переменная для операции по вычисению натурального логарифма

param = {
    'T': 'Т',
    'AO': 'АО',
    'abs_top': 'Абс. кровля репера (расчет)',
    'abs_bot': 'Абс. подошва репера (расчет)',

}


if __name__ == "__main__":
    # Read excel file with wells data
    excel_df = pd.read_excel(excel_wells_file, sheet_name=sheet_name)
    
    headers = excel_df.columns.values
    # print(headers)
     
    # Load LAS-file
    las = lasio.read('LAS+xlsx/320П.las', encoding='cp866')
    
    well_uwi = las.sections['Well']['UWI']['value']

    # Поиск нужных параметров в экселе по конкретному UWI скважины
    curr_well = excel_df.loc[excel_df['UWI'] == well_uwi, (param['T'], param['AO'], param['abs_top'] , param['abs_bot'])]
    
    # print(curr_well[param['abs_top']].iloc[0])
  

    well = las.df().reset_index()
    # print(well['DEPT'])


    depth_range = well.loc[(well['DEPT'] >= curr_well[param['abs_top']].iloc[0]) & (well['DEPT'] <= curr_well[param['abs_bot']].iloc[0])]
    for row in depth_range.index:
        print(depth_range['DEPT'][row], depth_range['GK'][row], depth_range['NKTB'][row])
 

"""
# print(curr_well['Т'].iloc[0])         # доступ к числовому значению нужного параметра из экселя




wb = load_workbook(filename='wells_data.xlsx', read_only=True)
ws = wb['Лист 1']


def get_headers():
    ''' Load headers from excel file'''
    headers = {}
    for row in ws.iter_rows(max_col=ws.max_column, max_row=1):
        for cell in row:       
            headers[cell.value] = cell.column
    return headers
    
       
"""