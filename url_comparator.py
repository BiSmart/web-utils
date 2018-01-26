#!/usr/bin/env python3

import re
import argparse

from openpyxl import load_workbook
import xlwt

import html as html_lib
import urllib.request
import urllib.error


TITLE_PATTERN = re.compile('<title.*?>(.*?)</title>')
H1_PATTERN = re.compile('<h1.*?>(?:<[A-Za-z]+?(?:\s.*?)?>)*(.*?)(?:</.+?>)*?</h1>')


def getPageData(url):
    data = {}

    try:
        resp = urllib.request.urlopen(url)
        html = html_lib.unescape(resp.read().decode('utf-8'))
    except urllib.error.URLError as e:
        data['error'] = 'url: {}'.format(e.reason)
    except urllib.error.HTTPError as e:
        data['error'] = 'http: {}'.format(e.code)
    except Exception:
        data['error'] = 'unknown error'
    else:
        title = TITLE_PATTERN.search(html)
        h1 = H1_PATTERN.search(html)
        data['title'] = title[1] if title else ''
        data['h1'] = h1[1] if h1 else ''

    return data


def compare_urls(data, other_url):
    results = []

    base_url = data[0]['url'].rstrip('/')

    path_pattern = re.compile('^{}(/.*)'.format(base_url))
    for obj in data:
        rel_path = path_pattern.search(obj['url'])[1]

        print('Читаем: {}'.format(rel_path))
        other_data = getPageData(other_url + rel_path)

        total = { 'url': rel_path }

        if 'error' in other_data:
            total['error'] = other_data['error']
        else:
            total['title'] = obj['title'] == other_data['title']
            total['h1'] = obj['h1'] == other_data['h1']

        results.append(total)

    return results


def parse_excel(path):
    result = []

    wb = load_workbook(path)
    sheet = wb.active

    for row in range(2, sheet.max_row+1):
        result.append({
            'url': sheet.cell(row=row, column=1).value,
            'title': sheet.cell(row=row, column=2).value,
            'h1': sheet.cell(row=row, column=3).value
        })

    return result


def output_excel(data):
    wb = xlwt.Workbook(encoding='utf-8')
    sheet = wb.add_sheet('Sheet 1')

    sheet.write(0, 0, 'URL')
    sheet.write(0, 1, 'TITLE')
    sheet.write(0, 2, 'H1')
    sheet.write(0, 3, 'ERROR')

    line_num = 1
    for record in data:
        if 'error' not in record and record['h1'] and record['title']:
            continue

        row = sheet.row(line_num)
        line_num += 1

        row.write(0, record['url'])
        if 'error' in record:
            row.write(3, record['error'])
        else:
            row.write(1, '' if record['title'] else 'Не совпадает')
            row.write(2, '' if record['h1'] else 'Не совпадает')

    wb.save('output.xls')


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument(
        'excel_path',
        help='Абсолютный или относительный путь до таблицы с данными')
    parser.add_argument(
        'mirror_path',
        help='базовый путь, с которым сравниваем урлы из таблицы')
    args = parser.parse_args()

    print('Читаем таблицу...')
    table = parse_excel(args.excel_path)

    print('Начинаем чтение url-ов...')
    offset = compare_urls(table, args.mirror_path)

    print('Сохраняем результат...')
    output_excel(offset)
