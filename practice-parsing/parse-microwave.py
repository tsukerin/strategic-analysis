import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from tqdm import tqdm
import re
import spacy

def sanitize_filename(name):
    return re.sub(r'[^\w-]', '_', name)

def get_channels():
    channels = []
    response = requests.get('https://www.microwavejournal.com/')
    if response.status_code != 200:
        raise Exception("Ошибка загрузки страницы")

    bs = BeautifulSoup(response.text, 'html.parser')
    for i in bs.find(class_='navigation').find_all(class_='level1-li'):
        if 'Channels' in i.find('a')['data-eventlabel']:
            num = 0
            for j in i.find(class_='level2').find_all(class_='level2-li'):
                href = j.find('a')['href']
                id_match = re.search(r'\d+', href)
                if id_match:
                    channels.append((num, j.find('a')['data-eventlabel'][6:j.find('a')['data-eventlabel'].find('|')], id_match.group()))
                    num += 1
    return channels

def get_max_pages(id):
    res = requests.get(f'https://www.microwavejournal.com/articles/topic/{id}?page=1')
    bs = BeautifulSoup(res.text, 'html.parser')
    return int(bs.find(class_='pagination').find_all('a')[-2].text)

def get_titles(bs):
    return bs.find_all(class_='headline article-summary__headline')

def get_date(bs):
    return bs.find_all(class_='date article-summary__post-date')

def get_full_text(bs):
    return bs.find_all(class_='abstract article-summary__teaser')

def get_links(bs):
    return bs.find_all(class_='headline article-summary__headline')

def extract_entities(text, nlp):
    doc = nlp(text)
    entities = [(ent.text, ent.label_) for ent in doc.ents]
    return entities

def parse(pages, id, name):
    nlp = spacy.load("en_core_web_sm")
    data = []

    for page in tqdm(range(1, pages + 1)):
        response = requests.get(f'https://www.microwavejournal.com/articles/topic/{id}?page={page}')
        if response.status_code != 200:
            print(f"Невозможно загрузить страницу {page}")
            continue

        bs = BeautifulSoup(response.text, 'html.parser')
        titles, dates, full_texts, links = get_titles(bs), get_date(bs), get_full_text(bs), get_links(bs)

        for i in range(len(titles)):
            title = titles[i].find('a').text
            date_text = dates[i].text
            link = links[i].find('a')['href']

            if i < len(full_texts):
                full_text_p = full_texts[i].find('p')
                if full_text_p:
                    full_text_content = full_text_p.text
                else:
                    full_text_content = "N/A"
            else:
                full_text_content = "N/A"

            if full_text_content == "N/A":
                entities = extract_entities(title, nlp)
            else:
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
    print('Выберите канал, статьи которого вы хотите спарсить:')

    channels = get_channels()
    
    for channel in channels:
        print(f'{channel[0]} - {channel[1]}')
    print(f'{max(channels, key=lambda x: x[0])[0]+1} - Спарсить все каналы в один файл Excel.')

    choice = int(input(f'Отправьте цифру от 0 до {channels[-1][0]+1}, чтобы выбрать желаемую опцию\n'))
    
    if choice in range(len(channels)):
        pages = int(input(f'Введите количество желаемых страниц, которые вы хотите спарсить. Максимум страниц - {get_max_pages(channels[choice][2])}:\n'))
   
    print('Пожалуйста, подождите...')

    if choice in range(len(channels)):
        parse(pages, channels[choice][2], channels[choice][1])
        print(f'Парсинг завершен! Файл {sanitize_filename(channels[choice][1])}.xlsx сохранен в директорию "output".')
    else:
        for channel in channels:
            print(f'Производится парсинг канала {channel[1]}')
            parse(int(get_max_pages(channel[2])), channel[2], 'microwave_articles')
        print(f'Парсинг завершен! Файл microwave_articles.xlsx сохранен в директорию "output".')
