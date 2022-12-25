import pandas as pd
import lasio, math, bcolors, re


# Название эксель файла со скважинами откуда считываются параметры
excel_wells_file = 'wells_data.xlsx'
sheet_name = 'Лист 1'

# Название файла куда будут записываться результаты
resulting_file_name = "s1-s2_toc_results.xlsx"

# Переменная для операции по вычисению натурального логарифма
ln = math.log1p                                 

# Необходимые мнемоники кривых которые понадобятся для расчетов
CURVES_MNEMONIC = ('GK', 'IK','NKT')

# Коэффициенты в формулах
coefficients = {

    'S1_S2': {
            'k': 7.7669,
            'k_t': -4.32833,
            'k_ao': 0.16643,
            'k_gk': 0.59841,
            'k_nkt': -4.64404,
            'k_ik': -4.2517,
             },

    'TOC': {
            'k': 8.712488,
            'k_t': -0.488716,
            'k_ao': 0.017437,
            'k_gk': 0.107496,
            'k_nkt': -0.945796,
            'k_ik': -0.746241,
            }

}

# Названия нужных столбцов в экселе из которых будем извлекать данные
param = {
            'T': 'Т',
            'AO': 'АО',
            'abs_top': 'Абс. кровля репера (расчет)',
            'abs_bot': 'Абс. подошва репера (расчет)',
        }



def curr_well_excel_data(las_uwi):
    '''Извлекаем из экселя необходимые данные по текущей скважине'''
    data = excel_df.loc[excel_df['UWI'] == las_uwi, (param['T'], param['AO'], param['abs_top'] , param['abs_bot'])]
    if data.empty:
        return print(bcolors.bcolors.FAIL + f'-----------\nLAS WELL UWI:{las_uwi} NOT FOUND IN EXCEL!!! \
        \n-----------\n' + bcolors.bcolors.ENDC)
    return data


def mnemonics_validate():
    ''' 
    Проверяем наличие мнемоник кривых в ласе с нашим необходимым для расчетов набором мнемоник.
    Исправляем имя кривой если оно немного отличается от имен в нашем наборе мнемоник.    
    '''
    count_curves = 0
    for curve in well_las:           
        for mnemonic in CURVES_MNEMONIC:
            if re.search('^' + mnemonic, curve):
                well_las.rename(columns={curve: mnemonic}, inplace=True)
                count_curves += 1
    if count_curves != 3:
        print(bcolors.bcolors.FAIL + f'-----------\nНе найдены некоторые кривые в лас файле скважины: {las_uwi}!!! \
        \n-----------\n' + bcolors.bcolors.ENDC)


def s1_s2(tempr, ao, gk, nkt, ik, **kwargs):
    '''Вычисление S1+S2'''
    k, k_t, k_ao, k_gk, k_nkt, k_ik = kwargs.values()
    # print(k, k_t, k_ao, k_gk, k_nkt, k_ik)

    # Строка с формулой
    result = k + k_t*tempr + k_ao*ao + k_gk*gk + k_nkt*nkt + k_ik*ln(ik) 
    return result


def toc(tempr, ao, gk, nkt, ik, **kwargs):
    '''Вычисление TOC'''
    k, k_t, k_ao, k_gk, k_nkt, k_ik = kwargs.values()
    # print(k, k_t, k_ao, k_gk, k_nkt, k_ik)

    # Строка с формулой
    result = k + k_t*tempr + k_ao*ao + k_gk*gk + k_nkt*nkt + k_ik*ln(ik) 
    return result



if __name__ == "__main__":
    # Read excel file with wells data
    excel_df = pd.read_excel(excel_wells_file, sheet_name=sheet_name)

    # Create resulting excel file
    file_result = pd.ExcelWriter(resulting_file_name)
      
    # Load LAS-file
    las = lasio.read('LAS+xlsx/320П.las', encoding='cp866')    
    las_uwi = las.sections['Well']['UWI']['value']

    # Поиск  нужных параметров в экселе по конкретному UWI скважины и их запись
    well_excel = curr_well_excel_data(las_uwi)    

    # Для удобства присвоим переменным значения нужных параметров из экселя со скважинами
    t, ao, abs_t, abs_b = (float(well_excel[i].to_numpy()) for i in well_excel)
    # print(t, ao, abs_t, abs_b)


    well_las = las.df().reset_index()
    # print(well['DEPT'])

    mnemonics_validate()

    depth_range_df = well_las.loc[(well_las['DEPT'] >= abs_t) & (well_las['DEPT'] <= abs_b)]
    # print(*depth_range_df.columns)

    gk = depth_range_df['GK']
    nkt = depth_range_df['NKT']
    ik = depth_range_df['IK']
    # print(nkt)
 
    for row in depth_range_df.index:
        s1_s2_result = s1_s2(t, ao, gk[row], nkt[row], ik[row], **coefficients['S1_S2'])
        toc_result = toc(t, ao, gk[row], nkt[row], ik[row], **coefficients['TOC'])
        print(s1_s2_result, toc_result)
 


    file_result.close()





"""
CURVES_MNEMONIC = {             
    'GK': {'GK'},
    'IK': {'IK'},
    'NKT': {'NKT', 'NKTB'},
    }



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