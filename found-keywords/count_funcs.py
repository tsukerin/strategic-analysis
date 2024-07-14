import os
import csv
from openpyxl import load_workbook

def find_keywords_by_function(directory, functions):
    keyword_counts = {func_name: 0 for func_name in functions}

    for filename in os.listdir(directory):
        if filename.endswith('.xlsx'):
            file_path = os.path.join(directory, filename)
            workbook = load_workbook(file_path)
            sheet = workbook.active

            for row in sheet.iter_rows(min_row=2, values_only=True):
                keywords = row[2]
                if keywords:
                    for keyword in keywords.split(', '):
                        for func_name, func_keywords in functions.items():
                            if any(partial_word.lower() in keyword.lower() for partial_word in func_keywords):
                                keyword_counts[func_name] += 1

    with open('found-keywords/function_keyword_counts.csv', 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(['Obj', 'Count'])
        for func_name, count in keyword_counts.items():
            writer.writerow([func_name, count])

if __name__ == '__main__':
    functions = {
        'Amplification': ['amplify', 'amplificate', 'amplification'],
        'Data Transmission': ['transmit', 'transmission', 'data transmission'],
        'Detection': ['detect', 'detection'],
        'Navigation': ['navigate', 'navigation'],
        'Communication': ['communicate', 'communication'],
        'Monitoring': ['monitor', 'monitoring'],
        'Identification': ['identify', 'identification'],
        'Real-time Communication': ['real-time', 'real time', 'real-time communication']
    }

    os.makedirs('found-keywords', exist_ok=True)
    find_keywords_by_function('practice-parsing/output', functions)
    print('Поиск ключевых слов завершен успешно! Результаты сохранены в found-keywords/function_keyword_counts.csv')
