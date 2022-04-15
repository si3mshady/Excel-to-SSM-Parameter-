import  boto3
import pandas as pd

EXCEL_SHEET = "server.xlsx"
ssm = boto3.client('ssm')

df = pd.ExcelFile(EXCEL_SHEET)


def get_sheet_count_and_names(df) -> tuple:
    return (len(df.sheet_names), df.sheet_names)

tab_count, tab_name_list = get_sheet_count_and_names(df)
print(tab_name_list[0])

sheet = df.parse(tab_name_list[0])
rows, columns = sheet.shape


def create_ssm_params(tab_list):    
    for tab in tab_list:
        sheet = df.parse(tab)
        for _, row in enumerate(sheet.values):
            key = row[0].split('.')
            key_clean = ''.join([k.strip().title() for k in key]) #split key by '.' uppercase each word and join back together value = row[1]
            value = row[1]
            for val in  key_clean.split(): #initially save as a string, must use split to create list 
                try:
                    if 'dev' in tab.lower():
                        print('dev')                                        
                        ssm_param = f'/Account/dev/{val}'
                        ssm.put_parameter(Name=ssm_param,Type='String', Value=value)
                        print(ssm_param, 'added!')
                except Exception as e:
                        if "ParameterAlreadyExists" in str(e):
                            print('parameter already exists')
                
                if 'qa' in tab.lower():
                    try:
                        ssm_param = f'/Account/qa/{val}'                                        
                        ssm.put_parameter(Name=ssm_param,Type='String', Value=value)
                        print(ssm_param, 'added!')
                    except Exception as e:
                        if "ParameterAlreadyExists" in str(e):
                            print('parameter already exists')
                    
                    
                if 'rod' in tab.lower():       
                    try:             
                        ssm_param = f'/Account/prod/{val}'
                        ssm.put_parameter(Name=ssm_param, Type='String', Value=value)
                        print(ssm_param, 'added!')
                    except Exception as e:
                        if "ParameterAlreadyExists" in str(e):
                            print('parameter already exists')


create_ssm_params(tab_name_list)