import requests  
from datetime import datetime  
import json
from pathlib import Path
import time
import pandas as pd
  

headers = {  
    "User-Agent": "Your User Agent",
    } 

def write_json(json_file: dict, file_name:str) -> None:
    '''
    Writes a dictionary to a JSON file within the '.files' directory.

    Args
    - json_file: The dictionary data to write to the file.
    - file_name: The name of the file (without extension) to create/save.

    return: None
    '''
    folder_path = '.files'
    Path(folder_path).mkdir(parents=True, exist_ok=True)
    with open(f'.files/{file_name}.json', 'w', encoding='utf-8') as file:
        json.dump(json_file, file, ensure_ascii=False, indent=4)

def get_vacancies_json(
        vacancie_name: str, 
        regions: list[int], 
        industry_id: list[int] = None
        ) -> list[dict]:
    '''
    Fetches vacancy data from the HeadHunter API based on search criteria.

    Args
    - vacancie_name: The job title or keyword to search for.
    - regions: A list of region IDs where the vacancies are located.
    - industry_id: An optional list of industry IDs to filter the vacancies.

    return: A list of dictionaries, each representing a vacancy item from the API response.
    '''
    url = "https://api.hh.ru/vacancies" 
    vacancies_json = []
    for i in range(0, 20):
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
    '''
    Fetches all geographical areas from the HeadHunter API and saves them to '.files/areas.json'.

    Args None

    return: None
    '''
    url = "https://api.hh.ru/areas" 
    response = requests.get(url, headers=headers) 
    areas = response.json()
    write_json(areas, 'areas')

def save_industries_as_json() -> None:
    '''
    Fetches all industry sectors from the HeadHunter API and saves them to '.files/industries.json'.

    Args None

    return: None
    '''
    url = "https://api.hh.ru/industries" 
    response = requests.get(url, headers=headers) 
    areas = response.json()
    write_json(areas, 'industries')


def save_vacancies_as_json(
        vacancie_name: str, 
        regions: list[int], 
        industry_id: list[int] = None
        ) -> None:
    '''
    Fetches vacancies based on the provided criteria and saves them to a JSON file named after the vacancy and the current date.

    Args
    - vacancie_name: The job title or keyword to search for.
    - regions: A list of region IDs where the vacancies are located.
    - industry_id: An optional list of industry IDs to filter the vacancies.

    return: None
    '''
    current_date_time = datetime.now().strftime("%Y-%m-%d")
    vacancies = get_vacancies_json(vacancie_name, regions, industry_id)
    file_name = f'{vacancie_name}_{current_date_time}'
    write_json(vacancies, file_name)

def get_vacancie_df(
        vacancy_name: str, 
        regions: list[int] = None, 
        industry_id: list[int] = None
        ) -> pd.DataFrame:
    '''
    Fetches vacancies and extracts specific details into a list of lists.

    Args
    - vacancy_name: The job title or keyword to search for.
    - regions: An optional list of region IDs where the vacancies are located.
    - industry_id: An optional list of industry IDs to filter the vacancies.

    return: A list of lists, where each inner list contains details for one vacancy: 
    [
        request, 
        date, 
        title, 
        area, 
        company, 
        url, 
        salary, 
        currency, 
        requirement, 
        responsibility
    ]
    '''
    data = []
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

            one_vacancy_dict = {
                "Запрос":vacancy_name,
                "Дата": current_date_time,
                "Вакансия": vacancy_title,
                "Город":area,
                "Компания":  company_name,
                "Ссылка":  vacancy_url,
                "ЗП":  salary,
                "Валюта":currency,
                "Требования": requirement,
                "Обязанности": responsibility
            }
            data.append(one_vacancy_dict)
    df = pd.DataFrame(data)
    return df

def get_vacancies_df_from_dict(
        vacancy_name_dict: dict[str:list[int]], 
        regions: list[int] = None
        ) -> pd.DataFrame:
    '''
    Aggregates vacancy data for multiple job titles/keywords provided in a dictionary

    Args
    - vacancy_name_dict: A dictionary mapping job titles/keywords to lists of industry IDs.
    - regions: An optional list of region IDs to apply to all searches.

    return: A combined list of lists containing vacancy details for all specified job titles/keywords.
    '''
    datadrames = []
    for vacancy_name in vacancy_name_dict.keys():
        datadrames.append(
            get_vacancie_df(
                vacancy_name, 
                regions, 
                vacancy_name_dict[vacancy_name]
                )
            )
        
    return pd.concat(datadrames, ignore_index=True)

