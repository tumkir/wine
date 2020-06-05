from http.server import HTTPServer, SimpleHTTPRequestHandler

from jinja2 import Environment, FileSystemLoader, select_autoescape
from datetime import datetime
import pandas
from collections import defaultdict

env = Environment(
    loader=FileSystemLoader('.'),
    autoescape=select_autoescape(['html', 'xml'])
)

template = env.get_template('template.html')

current_year = datetime.today().year
winery_age = current_year - 1920

excel_data = pandas.read_excel('./wine3.xlsx', keep_default_na=False).to_dict('records')

wines = defaultdict(list)

for wine in excel_data:
    wines[wine['Категория']].append(wine)


rendered_page = template.render(
    wines=wines
)

with open('index.html', 'w', encoding="utf8") as file:
    file.write(rendered_page)

server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
server.serve_forever()
