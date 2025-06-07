import requests  
import xlsxwriter  
import os  
from datetime import datetime  
import json
from pathlib import Path
import time
  

headers = {  
    "User-Agent": "Your User Agent",
    } 

def write_json(json_file: dict, file_name:str) -> None:
    folder_path = '.files'
    Path(folder_path).mkdir(parents=True, exist_ok=True)
    with open(f'.files/{file_name}.json', 'w', encoding='utf-8') as file:
        json.dump(json_file, file, ensure_ascii=False, indent=4)

def get_vacancies_json(vacancie_name: str, regions: list[int], industry_id: list[int] = None) -> list[dict]:
    url = "https://api.hh.ru/vacancies" 
    vacancies_json = []
    for i in range(0, 50):
        time.sleep(1)
        params = {  
            "text": vacancie_name,  
            "area": regions, 
            "per_page": 100,  
            'page': i,  
            'industry': industry_id,
        }
        response = requests.get(url, params=params, headers=headers)  
        if response.status_code == 200:  
            json_file = response.json().get("items", [])
            vacancies_json.extend(json_file)
        else:  
            print(f"Request @{vacancie_name} failed with status code: {response.status_code}. Response text: {response.text}")   
    return vacancies_json
        
def save_regions_as_json() -> None:
    """
    save all areas in .files/areas.json
    """
    url = "https://api.hh.ru/areas" 
    response = requests.get(url, headers=headers) 
    areas = response.json()
    write_json(areas, 'areas')

def save_industries_as_json() -> None:
    """
    save all industries in .files/areas.json
    """
    url = "https://api.hh.ru/industries" 
    response = requests.get(url, headers=headers) 
    areas = response.json()
    write_json(areas, 'industries')


def save_vacancies_as_json(vacancie_name: str, regions: list[int], industry_id: list[int] = None) -> None:
    """
    save vacancies in .files/areas.json
    """
    current_date_time = datetime.now().strftime("%Y-%m-%d")
    vacancies = get_vacancies_json(vacancie_name, regions, industry_id)
    file_name = f'{vacancie_name}_{current_date_time}'
    write_json(vacancies, vacancie_name)


def get_vacancies_list(vacancy_name: str, regions: list[int] = None, industry_id: list[int] = None) -> list[list[str]]:
    """
    return list, wich contains: 
    request, 
    current_date_time, 
    vacancy_title, 
    area, 
    company_name, 
    vacancy_url, 
    salary, 
    currency,
    requirement, 
    responsibility
    """
    output_data = []
    current_date_time = datetime.now().strftime("%Y-%m-%d")
    vacancies = get_vacancies_json(vacancy_name, regions, industry_id)
    if len(vacancies) > 0:  
        for vacancy in vacancies:  
            vacancy_title = vacancy.get("name")  
            vacancy_url = vacancy.get("alternate_url")  
            company_name = vacancy.get("employer", {}).get("name")  
            area = vacancy.get("area", {}).get("name") 
            requirement = vacancy.get("snippet", {}).get("requirement")
            responsibility = vacancy.get("snippet", {}).get("responsibility")
            salary = 0  
            currency = ''
            if vacancy.get("salary") is not None:  
                salary_from = vacancy.get("salary", {}).get("from")  
                salary_to = vacancy.get("salary", {}).get("to")  
                currency = vacancy.get("salary", {}).get("currency") 
                if salary_from is not None: salary = salary_from  
                if salary_to is not None: salary = salary_to   
            output_data.append([vacancy_name, current_date_time, vacancy_title, area, company_name, vacancy_url, salary, currency, requirement, responsibility])
    return output_data

def get_vacancies_list_from_dict(vacancy_name_dict: dict[str:list[int]], regions: list[int] = None) -> list[list[str]]:
    """
    collect all requests from get_vacancies_list() in one list
    """
    output_data = []
    for vacancy_name in vacancy_name_dict.keys():
        vacancies_list = get_vacancies_list(vacancy_name, regions, vacancy_name_dict[vacancy_name])
        output_data.extend(vacancies_list)
    return output_data

def save_xlsx(data: list[list], name: str, columns:list[dict]) -> None:
    """
    save data as xlsx in .files.
    """
    folder_path = '.files'
    Path(folder_path).mkdir(parents=True, exist_ok=True)
    dir_path = os.path.dirname(os.path.abspath(__file__))  
    current_date_time = datetime.now().strftime("%Y-%m-%d")
    workbook = xlsxwriter.Workbook(f'{dir_path}\\{folder_path}\\{name}_{current_date_time}.xlsx')  
    worksheet = workbook.add_worksheet()
    worksheet.add_table(  
        "A1:{0}{1}".format(get_column_letter(len(data[0])), len(data)+1),
        {  
            "data": data,
            "columns": columns,  
        }  
    )  
    workbook.close()

def get_column_letter(column_number:int) -> str:
    """
    1 -> 'A', 2 -> 'B', 27 -> 'AA'.
    """
    letter = ""
    while column_number > 0:
        column_number, remainder = divmod(column_number - 1, 26)
        letter = chr(65 + remainder) + letter
    return letter

columns = [  
    {"header": "Запрос"},
    {"header": "Дата"},
    {"header": "Вакансия"},  
    {"header": "Город"},  
    {"header": "Компания"},  
    {"header": "Ссылка"},  
    {"header": "ЗП"},  
    {"header": "Валюта"},
    {"header": "Требования"}, 
    {"header": "Обязанности"}, 
]
