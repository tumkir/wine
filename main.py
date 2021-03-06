import argparse
from collections import defaultdict
from datetime import datetime
from http.server import HTTPServer, SimpleHTTPRequestHandler

import pandas
from jinja2 import Environment, FileSystemLoader, select_autoescape

env = Environment(
    loader=FileSystemLoader('.'),
    autoescape=select_autoescape(['html', 'xml'])
)

template = env.get_template('template.html')


def parse_arguments():
    parser = argparse.ArgumentParser()
    parser.add_argument('filepath', help="Путь к .xlsx файлу с данными", default='wine.xlsx', nargs='?')
    args = parser.parse_args()
    return args


def calculate_winery_age():
    year_of_foundation = 1920
    current_year = datetime.today().year
    winery_age = current_year - year_of_foundation
    return winery_age


def read_data_from_excel_file(filepath):
    excel_data = pandas.read_excel(filepath, keep_default_na=False).to_dict('records')
    wines = defaultdict(list)
    for wine in excel_data:
        wines[wine['Категория']].append(wine)
    return wines


def main():
    filepath = parse_arguments().filepath

    try:
        wines_data = read_data_from_excel_file(filepath)
    except FileNotFoundError:
        print(f'Файл {filepath} не найден. При запуске скрипта укажите верный путь к файлу в аргументе командной строки.')
        return

    rendered_page = template.render(
        wines_data=wines_data.items(),
        winery_age=calculate_winery_age()
    )

    with open('index.html', 'w', encoding='utf8') as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
