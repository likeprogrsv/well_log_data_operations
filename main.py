import pandas as pd
import os, lasio, math, bcolors, re
from datetime import datetime


# Название эксель файла со скважинами откуда считываются параметры
excel_wells_file = 'wells_data_original.xlsx'
sheet_name = 'Лист 1'

# Путь к папке с ласами в корневой папке проекта 
las_files_path = 'LAS'

# Коэффициенты в формулах
coefficients = {

    'S1_S2': {
            'k': 7.7669,
            'k_t': 0,
            'k_ao': 0,
            'k_gk': 0.59841,
            'k_nkt': -4.64404,
            'k_ik': -4.2517,
             },

    'TOC': {
            'k': 8.712488,
            'k_t': 0,
            'k_ao': 0,
            'k_gk': 0.107496,
            'k_nkt': -0.945796,
            'k_ik': -0.746241,
            }
}

# Название файла куда будут записываться результаты и другие параметры файла
resulting_file_name = "s1-s2_toc_results.xlsx"
result_sheet1_name = 'sheet1'
result_sheet2_name = 'sheet2'
errors_sheet3_name = 'errors'
sheet1_columns = ['oil field', 'well', 'UWI', 'depth', 'T', 'AO', 'GK', 'NKT', 'IK', 'S1+S2', 'TOC']
sheet2_columns = ['oil field', 'well', 'UWI', 'depth top', 'depth bot', 'T', 'AO', 'GK mean', 'NKT mean', 'IK mean', \
    'S1+S2 mean', 'TOC mean']

# Переменная для операции по вычисению натурального логарифма
ln = math.log1p                                 

# Необходимые мнемоники кривых которые понадобятся для расчетов
CURVES_MNEMONIC = ('GK', 'IK','NKT')

# Скважины с ошибками в кривых (отрицательные значения). по скважинам для кривой было присвоено значение 0.2 
curves_errors = {}

# Названия нужных столбцов в экселе из которых будем извлекать данные
param = {
            'oil_field': 'Месторождение',
            'well_name': 'Скважина',
            'T': 'Т',
            'AO': 'АО',
            'depth_top': 'Кровля репера',
            'depth_bot': 'Подошва репера',
            'abs_top': 'Абс. кровля репера (расчет)',
            'abs_bot': 'Абс. подошва репера (расчет)',
        }

# Путь к корневой папке проекта. НЕ ИЗМЕНЯТЬ
project_path = os.path.abspath('.')



def curr_well_excel_data(las_uwi):
    '''Извлекаем из экселя необходимые данные по текущей скважине'''
    data = excel_df.loc[excel_df['UWI'] == las_uwi, (param['oil_field'], param['well_name'], param['T'], \
        param['depth_top'], param['depth_bot'], param['AO'], param['abs_top'] , param['abs_bot'])]
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


def validate_curve_value(well_name, gk, nkt, ik):
    
    d = {'gk': gk, 'nkt': nkt, 'ik': ik}
    for i in d:
        if -999 < d[i] <= 0:
            d[i] = 0.2
            curves_errors[well_name] = True
    
    return d


def s1_s2(well_name, tempr, ao, gk, nkt, ik, **kwargs):
    '''Вычисление S1+S2'''
    k, k_t, k_ao, k_gk, k_nkt, k_ik = kwargs.values()
    d = validate_curve_value(well_name, gk, nkt, ik)
    # Строка с формулой
    result = k + k_t*tempr + k_ao*ao + k_gk*d['gk'] + k_nkt*d['nkt'] + k_ik*ln(d['ik']) 
    return result


def toc(well_name, tempr, ao, gk, nkt, ik, **kwargs):
    '''Вычисление TOC'''
    k, k_t, k_ao, k_gk, k_nkt, k_ik = kwargs.values()
    d = validate_curve_value(well_name, gk, nkt, ik)
    # Строка с формулой
    result = k + k_t*tempr + k_ao*ao + k_gk*d['gk'] + k_nkt*d['nkt'] + k_ik*ln(d['ik']) 
    return result


def recreate_resulting_file():
    '''Пересоздаём итоговый файл для записи новых данных'''
    file_result = pd.ExcelWriter(resulting_file_name)
    [file_result.book.create_sheet(i) for i in [result_sheet1_name, result_sheet2_name, errors_sheet3_name]]
    file_result.book.save(resulting_file_name)
    file_result.book.close()


def write_results(writer, sheet_name, dataframe):
    '''Записываем в отчетный файл необходимую инфу'''    
    if writer.sheets[sheet_name].max_row > 1:
        dataframe.to_excel(writer, sheet_name=sheet_name, startrow=writer.sheets[sheet_name].max_row, \
        header=False, index=False)
    else:
        dataframe.to_excel(writer, sheet_name=sheet_name, index=False)
    


if __name__ == "__main__":
    start = datetime.now()
    # Read excel file with wells data
    excel_df = pd.read_excel(excel_wells_file, sheet_name=sheet_name)

    # Пересоздаём итоговый файл для записи новых данных
    recreate_resulting_file()

    '''
    ------------------------------------------------------------------------------- 
    Проходим в цикле по всем ласам в папке, вычисляем параметры и записываем в файл
    -------------------------------------------------------------------------------
    '''

    for filename in os.scandir(project_path + '\\' +las_files_path):
        if filename.is_file():         
            print(f'Working on {filename.name}')
            path = las_files_path + '/' + filename.name
            # Load LAS-file
            try:
                las = lasio.read(las_files_path + '/' + filename.name, encoding='cp866')    
            except ValueError:
                print(bcolors.bcolors.FAIL + f'-----------\nЧто-то не так с данными кривых в ласе скважины: {filename.name}!!! \
                    \n-----------\n' + bcolors.bcolors.ENDC)
                continue            
            las_uwi = las.sections['Well']['UWI']['value']

            # Поиск  нужных параметров в экселе по конкретному UWI скважины и их запись
            well_excel = curr_well_excel_data(las_uwi)    

            # Для удобства присвоим переменным значения нужных параметров из экселя со скважинами
            try:
                oil_field, well_name, t, depth_t, depth_b, ao, abs_t, abs_b = (float(well_excel[i].to_numpy()) if well_excel[i].dtype != \
                    'object' else well_excel[i].values.astype(str)[0] for i in well_excel)
            except TypeError:                
                with pd.ExcelWriter(resulting_file_name, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:                
                    errors_df = pd.DataFrame({'Скв отсутст в экселе': filename.name, 'Скв с отриц знач в кривых': '-'}, \
                        index=[writer.sheets[errors_sheet3_name].max_row])
                    write_results(writer, errors_sheet3_name, errors_df)
                continue

            well_las = las.df().reset_index()

            mnemonics_validate()

            depth_range_df = well_las.loc[(well_las['DEPT'] >= depth_t) & (well_las['DEPT'] <= depth_b)]

            gk = depth_range_df['GK']
            nkt = depth_range_df['NKT']
            ik = depth_range_df['IK']
            
            # Creating dataframes for saving in excel
            sheet1_df = pd.DataFrame(columns=sheet1_columns)
            sheet2_df = pd.DataFrame(columns=sheet2_columns)
           
            for row in depth_range_df.index:
                s1_s2_result = s1_s2(filename.name, t, ao, gk[row], nkt[row], ik[row], **coefficients['S1_S2'])
                toc_result = toc(filename.name, t, ao, gk[row], nkt[row], ik[row], **coefficients['TOC'])
                new_row = pd.DataFrame([oil_field, well_name, las_uwi, depth_range_df['DEPT'][row], t, ao, gk[row], nkt[row], \
                    ik[row], s1_s2_result, toc_result], index=sheet1_columns).T                
                sheet1_df = pd.concat([sheet1_df, new_row], ignore_index=True)

            row_for_sheet2 = pd.DataFrame([oil_field, well_name, las_uwi, depth_t, depth_b, t, ao, sheet1_df['GK'].mean(), \
                sheet1_df['NKT'].mean(), sheet1_df['IK'].mean(), sheet1_df['S1+S2'].mean(), sheet1_df['TOC'].mean()], index=sheet2_columns).T
            sheet2_df = pd.concat([sheet2_df, row_for_sheet2], ignore_index=True)

            with pd.ExcelWriter(resulting_file_name, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                # Записываем первую книгу отчета со всеми строками данных
                write_results(writer, result_sheet1_name, sheet1_df)
                
                # Записываем вторую сводную книгу отчета со средними рассчитанными значениями по каждой скважине
                write_results(writer, result_sheet2_name, sheet2_df)

                # Записываем ошибки
                if filename.name in curves_errors:   
                    errors_df = pd.DataFrame({'Скв отсутст в экселе': '-', 'Скв с отриц знач в кривых': well_name}, \
                        index=[writer.sheets[errors_sheet3_name].max_row])
                    write_results(writer, errors_sheet3_name, errors_df)    
    
    total_time = datetime.now() - start
    print(bcolors.bcolors.OKBLUE + f'Время работы программы: {total_time}' + bcolors.bcolors.ENDC)


'''
            with pd.ExcelWriter(resulting_file_name, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                # Записываем первую книгу отчета со всеми строками данных
                if writer.sheets[result_sheet1_name].max_row > 1:
                    sheet1_df.to_excel(writer, sheet_name=result_sheet1_name, startrow=writer.sheets[result_sheet1_name].max_row, \
                        header=False, index=False)           
                else:            
                    sheet1_df.to_excel(writer, sheet_name=result_sheet1_name, index=False)
                
                # Записываем вторую сводную книгу отчета со средними рассчитанными значениями по каждой скважине
                if writer.sheets[result_sheet2_name].max_row > 1:  
                    sheet2_df.to_excel(writer, sheet_name=result_sheet2_name, startrow=writer.sheets[result_sheet2_name].max_row, \
                        header=False, index=False)                          
                else:
                    sheet2_df.to_excel(writer, sheet_name=result_sheet2_name, index=False)
'''
    # errors_df = pd.DataFrame({'Скв отсутст в экселе': well_not_found, 'Скв с отриц знач в кривых': curves_errors})
    # with pd.ExcelWriter(resulting_file_name, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
    #     errors_df.to_excel(writer, sheet_name=errors_sheet3_name, index=False)



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