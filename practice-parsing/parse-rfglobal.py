import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from tqdm import tqdm
import re
import spacy

def sanitize_filename(name):
    return re.sub(r'[^\w-]', '_', name)

def get_titles(bs):
    return bs.find_all('a', class_='d-block')

def get_date(bs):
    return bs.find_all('em', class_='vm-hub-date')

def get_full_text(bs):
    return bs.find_all('div', class_='pt-1')

def extract_entities(text, nlp):
    doc = nlp(text)
    entities = [(ent.text, ent.label_) for ent in doc.ents]
    return entities

def parse(pages, name):
    nlp = spacy.load("en_core_web_sm")
    data = []

    for page in tqdm(range(1, pages + 1)):
        response = requests.get(f'https://www.rfglobalnet.com/hub/bucket/feature-articles?page={page}')
        if response.status_code != 200:
            print(f"Невозможно загрузить страницу {page}")
            continue

        bs = BeautifulSoup(response.text, 'html.parser')
        titles, dates, full_texts = get_titles(bs), get_date(bs), get_full_text(bs)

        for i in range(len(titles)):
            title = titles[i].text
            date_text = dates[i].text
            link = 'https://www.rfglobalnet.com' + titles[i]['href']

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
            data.append([title, date_text, full_text_content, entity_text, link])

    sanitized_name = sanitize_filename(name)
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(['Title', 'Date', 'Full Text', 'Entities', 'Link'])
    
    for row in data:
        sheet.append(row)

    workbook.save(f'practice-parsing/output/{sanitized_name}.xlsx')

if __name__ == '__main__':
    print('Пожалуйста, подождите...')
    parse(10, 'rfglobalnet')
    print(f'Парсинг завершен! Файл rfglobalnet.xlsx сохранен в директорию "output".')
