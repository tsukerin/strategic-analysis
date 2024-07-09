import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from tqdm import tqdm
import re
import spacy

def sanitize_filename(name):
    return re.sub(r'[^\w-]', '_', name)

def get_titles(bs):
    return bs.find_all('div', class_='auswahl_beitrag')  

def get_full_text(bs):
    return bs.find_all('div', class_='excerpt-div')

def extract_entities(text, nlp):
    doc = nlp(text)
    entities = [(ent.text, ent.label_) for ent in doc.ents]
    return entities

def parse(pages, name):
    nlp = spacy.load("en_core_web_sm")
    data = []

    for page in tqdm(range(1, pages + 1)):
        response = requests.get(f'https://euro-sd.com/archive/?_sft_category=technology&sf_paged={page}')
        if response.status_code != 200:
            print(f"Невозможно загрузить страницу {page}")
            continue

        bs = BeautifulSoup(response.text, 'html.parser')
        titles, full_texts = get_titles(bs), get_full_text(bs)

        for i in range(len(titles)):
            title = titles[i].find('a')['title']
            link = titles[i].find('a')['href']

            if i < len(full_texts):
                full_text_p = full_texts[i].find('p')
                if full_text_p:
                    full_text_content = full_text_p.text
                else:
                    full_text_content = "N/A"
            else:
                full_text_content = "N/A"

            entities = extract_entities(full_text_content, nlp)
            entity_text = ", ".join([f"{ent[0]} ({ent[1]})" for ent in entities])
            data.append([title, full_text_content, entity_text, link])

    sanitized_name = sanitize_filename(name)
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(['Title', 'Full Text', 'Entities', 'Link'])
    
    for row in data:
        sheet.append(row)

    workbook.save(f'practice-parsing/output/{sanitized_name}.xlsx')

if __name__ == '__main__':
    print('Пожалуйста, подождите...')
    parse(16, 'parse-euro-sd')
    print(f'Парсинг завершен! Файл parse-euro-sd.xlsx сохранен в директорию "output".')
