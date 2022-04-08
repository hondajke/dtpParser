import requests
from fake_useragent import UserAgent
import json
import openpyxl
import os.path

def get_vins():
    vins = []
    with open('input.txt', encoding='utf-8') as f:   
        while True:
            term = f.readline()
            if not term:
                break
            quer = term.strip()
            vins.append(quer)
            #print(quer)
    return vins

def post_vins(data, errorcount: int, emptycount: int, timeouttime: int):
    key = 2
    fileId = 1
    empty = 0
    try:
        wb = openpyxl.Workbook()
        worksheet = wb['Sheet'] #Делаем его активным
        worksheet['A1']='VIN'
        worksheet['B1']='Количество происшествий'
        worksheet['C1']='Инф о происшествии'
        worksheet['D1']='Дата'
        worksheet['E1']='Время'
        worksheet['F1']='Тип происшествия'
        worksheet['G1']='Регион'
        worksheet['H1']='Марка(модель)'
        worksheet['I1']='Год выпуска ТС'
        worksheet['J1']='Номер ТС'
        worksheet['K1']='из всего ТС в ДТП'
        #В указанную ячейку на активном листе пишем все, что в кавычках
        file_path = f'test{fileId}.xlsx'
        while os.path.exists(f'test{fileId}.xlsx'):
            fileId = fileId + 1
        wb.save(f'test{fileId}.xlsx')
    except:
        print('Ошибка при создании файла')
    
    for i in range(len(data)):
        if empty >= emptycount:
            break
        url = 'https://xn--b1afk4ade.xn--90adear.xn--p1ai/proxy/check/auto/dtp'
        postData = {
            'vin': data[i],
            'checkType': 'aiusdtp'
        }
        flag = False
        for j in range(errorcount):
            headers = {'User-Agent': UserAgent().random }
            try:
                r = requests.post(url, data=postData, headers=headers, timeout=timeouttime*(j + 1))
            except:
                print('timeout')
                continue
            print(r.status_code)
            if r.status_code == 200:
                #print(r.json())
                flag=True
                result = json.loads(r.text)
                accidents = result['RequestResult']['Accidents']
                print(len(accidents))
                if len(accidents) == 0:
                    try:
                        worksheet[f'A{key}']=data[i]
                        worksheet[f'B{key}']=0
                        #empty = empty + 1
                        key = key + 1
                        wb.save(f'test{fileId}.xlsx')  
                    except:
                        print('Ошибка при записи')
                else:
                    for k in range(len(accidents)):
                        try:
                            worksheet[f'A{key}']=data[i]
                            worksheet[f'B{key}']=len(accidents)
                            worksheet[f'C{key}']='№' + accidents[k]['AccidentNumber']
                            date = accidents[k]['AccidentDateTime']
                            date = date.split(' ')
                            worksheet[f'D{key}']=date[0]
                            worksheet[f'E{key}']=date[1]
                            worksheet[f'F{key}']=accidents[k]['AccidentType']
                            worksheet[f'G{key}']=accidents[k]['RegionName']
                            worksheet[f'H{key}']=accidents[k]['VehicleMark'] + ' ' + accidents[k]['VehicleModel']
                            worksheet[f'I{key}']=accidents[k]['VehicleYear']
                            worksheet[f'J{key}']=accidents[k]['VehicleSort']
                            worksheet[f'K{key}']=accidents[k]['VehicleAmount']
                            empty = 0
                            key = key + 1
                            wb.save(f'test{fileId}.xlsx')  
                        except:
                            print('Ошибка при записи')  
                break
            #elif r.status_code == 404:
            #    try:
            #        worksheet[f'A{key}']=data[i]
            #        worksheet[f'B{key}']='VIN не обнаружен'
            #        empty = empty + 1
            #        key = key + 1
            #        wb.save(f'test{fileId}.xlsx')  
            #    except:
            #        print('Ошибка при записи')  
            #    print('Not Found')
            #    break
        if flag == False:
            print('error')
            try:
                worksheet[f'A{key}']=data[i]
                worksheet[f'B{key}']='error'
                key = key + 1
                wb.save(f'test{fileId}.xlsx')  
            except:
                print('Ошибка при записи')  
        
            
            

if __name__ == '__main__':
    errcnt = 3
    with open('errorcount.txt', encoding='utf-8') as f:
        try:
            errorcount = f.readline()
            errcnt = int(errorcount)
        except:
            errcnt = 3
    
    emptycnt = 200
    with open('emptycount.txt', encoding='utf-8') as f:
        try:
            emptycount = f.readline()
            emptycnt = int(emptycount)
        except:
            emptycnt = 3
    
    timeouttime = 3
    with open('timeout.txt', encoding='utf-8') as f:
        try:
            tmt = f.readline()
            timeouttime = int(tmt)
        except:
            timeouttime = 3
            
    print('Количество запросов: ' + str(errcnt))
    print('Количество пустых запросов: ' + str(emptycnt))
    print('Время тайм-аута на один запрос: ' + str(timeouttime))
    vins = get_vins()
    post_vins(vins, errcnt, emptycnt, timeouttime)