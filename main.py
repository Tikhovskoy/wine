import os
import argparse
import pandas as pd
from jinja2 import Environment, FileSystemLoader
from http.server import HTTPServer, SimpleHTTPRequestHandler
from collections import defaultdict
import math
from datetime import datetime

def get_year_word(number: int) -> str:
    """Возвращает правильную форму слова 'год' в зависимости от числа."""
    if 11 <= number % 100 <= 19:
        return "лет"
    last_digit = number % 10
    if last_digit == 1:
        return "год"
    elif 2 <= last_digit <= 4:
        return "года"
    return "лет"

def calculate_winery_age(foundation_year=1920):
    """Возвращает возраст винодельни."""
    current_year = datetime.now().year
    age = current_year - foundation_year
    return age

def load_wine_data(file_path):
    """Загружает и группирует данные о вине из Excel."""
    df = pd.read_excel(file_path, engine="openpyxl", na_values=[""], keep_default_na=False)
    df = df.where(pd.notna(df), None)  
    grouped_products = defaultdict(list)
    for _, row in df.iterrows():
        category = row["Категория"]
        product = row.drop("Категория").to_dict()
        grouped_products[category].append(product)
    return grouped_products

def render_template(grouped_products, age, age_word):
    """Генерирует HTML-файл на основе шаблона."""
    env = Environment(loader=FileSystemLoader('.'))
    template = env.get_template('template.html')
    rendered_html = template.render(
        grouped_products=grouped_products,
        age=age,
        age_word=age_word
    )
    with open("index.html", "w", encoding="utf-8") as f:
        f.write(rendered_html)

def start_server():
    """Запускает локальный сервер для отображения сайта."""
    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()

def parse_arguments():
    """Парсит аргументы командной строки."""
    parser = argparse.ArgumentParser(description="Запуск винного каталога")
    parser.add_argument("--data", default=os.getenv("WINE_DATA_FILE", "wine_catalog.xlsx"),
                        help="Путь к файлу с данными (по умолчанию wine_catalog.xlsx)")
    return parser.parse_args()

def main():
    """Основная логика программы."""
    args = parse_arguments()
    age = calculate_winery_age()
    age_word = get_year_word(age)
    grouped_products = load_wine_data(args.data)
    render_template(grouped_products, age, age_word)
    start_server()

if __name__ == "__main__":
    main()
