import sys, time

try:
    import xlrd
except Exeption as err:
    print (ret_all(err))
    exit()
    
    
def get_sheet_names(file_path):
    '''
    Get info about all sheet names
    :return: dict
    '''
    result = False
    start = time.time()
    
    wb = None
    try:
        wb = xlrd.open_workbook(file_path)
    except Exception as err:
        names = ret_all(err)
    if wb:
        try:
            names =  wb.sheet_names()
            result = True
        except Exception as err:
            names = ret_all(err)
        
    stop = time.time()
    timings = str(round(stop - start, 3)) + 's'

    return {'result': result,
            'summary': names,
            'timing': timings}  


def get_sheet_headers(file_path, sheet_name, row_offset=0):
    '''
    Get all headers including empty cells
    :return: dict
    '''

    result = False
    start = time.time()
    headers = wb = None
    cols = 0 # initial number of cols
    try:
        wb = xlrd.open_workbook(file_path)
    except Exception as err:
        headers = ret_all(err)
    if wb:
        try:
            ws = wb.sheet_by_name(sheet_name)
            tst = ws.row_values(row_offset)
            headers = [str(r) for r in tst ] # convert all values to str type
            cols = len(headers)
            result = True
        except Exception as err:
            headers = ret_all(err)
         
    stop = time.time()
    timings = str(round(stop - start, 3)) + 's'

    return {'result': result,
            'summary': headers,
            'timing': timings, 
            'columns': cols}
            
def ret_all(err):
    return 'Error on line {!s}: {!s}: {!s}'.format(sys.exc_info()[-1].tb_lineno, type(err).__name__, err)
    
if __name__ == '__main__':
    print(get_sheet_names('test.xlsx'))
    print(get_sheet_headers('test.xlsx','main_sheet'))
    
    