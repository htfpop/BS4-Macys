""" **************************************************************
* Programmer : Chris Leung                                       *
* Project    : Macy's Parser                                     *
* Purpose    : Scrape all Macy's men's clothes items and sort by *
*              number of reviews in descending order             *
*****************************************************************"""

from bs4 import BeautifulSoup
from tabulate import tabulate
from openpyxl import Workbook
import requests

PARENT_WEBSITE = 'https://www.macys.com/shop/mens-clothing/all-mens-clothing/Pageindex/'
PARENT_SUFFIX = '?id=197651'
WEBSITE_PREFIX = "https://www.macys.com"
PAGES = 0
headers = {"User-Agent": "Mozilla/5.0"}
data = [
    ['Item', 'Brand', 'Reviews', 'Link']
]

def main():
    global PAGES
    for i in range(1, PAGES+1):  # PAGES = 1072
        render = PARENT_WEBSITE + str(i) + PARENT_SUFFIX
        # DEBUG
        print(f'\nRender: {render}\n')
        html = requests.get(render, headers=headers).text
        soup = BeautifulSoup(html, 'html.parser')
        dsx = soup.find_all(class_='productDescription')

        for div in dsx:
            link = div.find('a')
            website = WEBSITE_PREFIX + link.get('href')
            title = link.get('title')
            try:
                reviews = div.find('div', class_='stars').find('span', class_='aggregateCount').text.strip('()')
            except AttributeError:
                reviews = 'N/A'

            brand = div.find('div', class_='productBrand').text.strip()

            data.append([title, brand, reviews, website])

            # DEBUG
            # print(f'Item Title:{title:30} Website: {WEBSITE_PREFIX+website:35} Stars: {stars:3}')

    # Perform sorting based off of reviews. If aggregateCount cannot be put as an integer, assign it negative infinity
    # to put at the bottom of the list. Highest amount of reviews first
    sorted_tab = sorted(data[2:], key=lambda x: int(x[2]) if x[2].isdigit() else float('-inf'), reverse=True)

    # Append column titles to the very first item in the dictionary
    sorted_tab.insert(0, data[0])

    # Create table and print
    print(tabulate(sorted_tab, headers="firstrow"))

    workbook = Workbook()
    sheet = workbook.active

    for row in sorted_tab:
        if row[2].isdigit():
            row[2] = int(row[2])
        sheet.append(row)

    workbook.save('test.xlsx')


def pre_check():
    try:
        import bs4
    except ImportError:
        print('[DEPENDENCY]: Could not import BeautifulSoup4. Please run \'pip install bs4\'')

    try:
        import tabulate
    except ImportError:
        print('[DEPENDENCY]: Could not import tabulate. Please run \'pip install tabulate\'')

    try:
        import openpyxl
    except ImportError:
        print('[DEPENDENCY]: Could not import tabulate. Please run \'pip install openpyxl\'')


def find_pages():
    global PAGES
    render = PARENT_WEBSITE + str(1) + PARENT_SUFFIX
    html = requests.get(render, headers=headers).text
    soup = BeautifulSoup(html, 'html.parser')

    index = soup.find('ul', class_='pagination').find('option').text
    PAGES = int(index.split("of")[1])
    print(PAGES)


if __name__ == '__main__':
    pre_check()
    find_pages()
    main()
