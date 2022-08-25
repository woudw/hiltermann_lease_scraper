import logging
from bs4 import BeautifulSoup
import pprint
import xlsxwriter
import requests


def get_content() -> str:
    """
    Get all info from Hiltermann
    """
    logging.info("Getting data from Hiltermann")
    url = "https://auto.hiltermannlease.nl/configurator/ajax/zoek/uitvoering-list/"

    headers = {
    'Accept': '*/*',
    'Accept-Language': 'en-US,en;q=0.5',
    'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
    'X-Requested-With': 'XMLHttpRequest',
    'Origin': 'https://auto.hiltermannlease.nl',
    'Connection': 'keep-alive',
    'Referer': 'https://auto.hiltermannlease.nl/configurator/zoek/overzicht/',
    'Sec-Fetch-Dest': 'empty',
    'Sec-Fetch-Mode': 'cors',
    'Sec-Fetch-Site': 'same-origin'
    }

    response = requests.request("POST", url, headers=headers)
    return response.text


def parse_web_result(soup: BeautifulSoup) -> list:
    """
    Converts HTML soup to a list of tuples (aka, a table)
    """
    logging.info('parsing web result')
    result = []
    for uitvoering_groep in soup.find_all('div', class_ = 'uitvoering-group'):
        uitvoering_groep_text = uitvoering_groep.find('h2').text
        table = uitvoering_groep.find('table')
        uitvoeringen = table.find_all('tr', class_ = 'uitvoering')
        for uitvoering in uitvoeringen:
            car_rec = {}
            car_rec['uitvoering_groep'] = uitvoering_groep_text
            a = uitvoering.attrs
            for k, v in uitvoering.attrs.items():
                if k != 'class':
                    # logging.debug(f"{k}, {v}")
                    car_rec[k] = v
        
            uitvoering_cell = uitvoering.find('td', class_ = 'uitvoering-cell')

            info_link = uitvoering_cell.find('a', class_ = 'info-link')
            info_link_href = info_link['href']
            car_rec['info-link'] = info_link_href

            uitvoering_link = uitvoering.find('a', class_ = 'uitvoering-link')
            uitvoering_link_href = uitvoering_link['href']
            car_rec['uitvoering-link'] = uitvoering_link_href

            for span in uitvoering_link.find_all('span'):
                key = span['class'][0]
                value = span.text
                car_rec[key] = value
                # logging.debug(f"{key} - {value}")

            prijs_cell = uitvoering.find('td', class_ = 'prijs-cell')
            prijs = prijs_cell.text
            car_rec['prijs'] = prijs

            lease_prijs_cell = uitvoering.find('td', class_ = 'lease-prijs-cell')
            lease_prijs = lease_prijs_cell.text
            car_rec['lease-prijs'] = lease_prijs

            result.append(car_rec)
            # end verwerk uitvoering
        # end loop uitvoeringen
    # end loop uitvoeringen-groep
    logging.info(f"Num of records: {len(result)}")
    pp = pprint.PrettyPrinter(indent=4)
    # pp.pprint(result)
    return result


def write_excel(data: list):
    logging.info('writing excel file')
    workbook = xlsxwriter.Workbook('autos.xlsx')
    worksheet = workbook.add_worksheet()

    columns = []
    for dict in data:
        for k in dict:
            if k not in columns:
                columns.append(k)
    for c in columns:
        i = columns.index(c)
        worksheet.write(0, i, c)

    row = 1
    for rec in data:
        for k, v in rec.items():
            i = columns.index(k)
            worksheet.write(row, i, v)
        row += 1
    workbook.close()


def main():
    logging.basicConfig(format='%(levelname)s:%(message)s', level=logging.DEBUG)
    content = get_content()
    soup = BeautifulSoup(content, 'html.parser')
    data = parse_web_result(soup)
    write_excel(data)


if __name__ == '__main__':
    main() 