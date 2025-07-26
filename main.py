import hh_requests as hh




# hh.save_industries_as_json() 
# hh.save_regions_as_json()
# hh.save_vacancies_as_json("bim-manager", [1, 2])
# hh.get_vacancies_list("bim-manager", None)
vacancies_dict = {
    'bim-менеджер':None, 
    'архитектор':[13], 
    'инженер-конструктор':[13], 
    'инженер ПТО':[13], 
    'инженер ОВиК':[13], 
    'инженер ВК':[13], 
    'Главный инженер проекта':[13],
    'курьер': None}
data = hh.get_vacancies_list_from_dict(vacancies_dict, [1, 2])
hh.save_xlsx(data, 'Вакансии')