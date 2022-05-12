from random import randrange
from time import sleep
from bs4 import BeautifulSoup
import requests
from fake_useragent import UserAgent
import xlsxwriter


links = []
data = {}
ua = UserAgent(verify_ssl=False)
headers = {
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
    'User-Agent': ua.random
}


def get_links(url):
    response = requests.get(url=url, headers=headers)
    with open(f'index.html', 'w') as file:
        file.write(response.text)

    with open('index.html') as file:
        src = file.read()

    soup = BeautifulSoup(src, 'lxml')

    for i in soup.find_all('li', class_='entry'):
        link = i.find('a', class_='woocommerce-LoopProduct-link').get('href')
        links.append(link)


def get_data(links):
    row, col = 1, 0

    workbook = xlsxwriter.Workbook('export.xlsx')
    worksheet = workbook.add_worksheet()

    worksheet.write('A1', 'URL')
    worksheet.write('B1', 'Артикул')
    worksheet.write('C1', 'Категория')
    worksheet.write('D1', 'Товар')
    worksheet.write('E1', 'Цена')
    worksheet.write('F1', 'Описание')

    count = 1
    print(f'[START]...Количество товаров: {len(links)}')

    for link in links:
        response = requests.get(url=link, headers=headers)
        soup = BeautifulSoup(response.text, 'lxml')
        title = soup.find('h2', class_='single-post-title').text
        price = soup.find('p', class_='price').text.split()[0]
        try:
            vendor = soup.find('span', class_='sku').text
        except AttributeError:
            vendor = 'Нет артикла'
        category = soup.find('span', 'posted_in').find('a').text
        try:
            description = soup.find('div', class_='woocommerce-Tabs-panel').find('p').text
        except AttributeError:
            description = 'Нет описания'

        worksheet.write(row, 0, link)
        worksheet.write(row, 1, vendor)
        worksheet.write(row, 2, category)
        worksheet.write(row, 3, title)
        worksheet.write(row, 4, price)
        worksheet.write(row, 5, description)
        row += 1
        print(f'[INFO]...{count}/{len(links)}.....{title}')
        count += 1
        sleep(randrange(0, 2))

        # print(vendor, title, price, category)
        # print(description)

    workbook.close()
    print(f'[DONE]...Работа завершена')


def main():
    get_links('https://optexpress.ru/shop/?products-per-page=all')
    get_data(links)


if __name__ == '__main__':
    main()

# di = {
#     'title': {
#         'url': '',
#         'price': '',
#         'category': '',
#         'description': '',
#         'vendor': '',
#
#     }
# }