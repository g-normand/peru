from bs4 import BeautifulSoup
import requests
from datetime import datetime
import re
import xlrd
from xlutils.copy import copy as xl_copy



def normalize(name):
    """Make names comparable regardless of hyphens/case/extra spaces."""
    name = name.strip().lower()
    name = name.replace('-', ' ').replace('’', "'")
    name = re.sub(r'\s+', ' ', name)
    return name


out_xls = 'lista_ok.xls'
all_species_in_country = True
bird_name_in_english = False
if bird_name_in_english:
    print('SEARCHING FOR ENGLISH BIRDNAME')
else:
    print('SEARCHING FOR FRENCH BIRDNAME')

if (all_species_in_country == True):
    # Get all Peru birds
    regions = {
        'PE': 'https://ebird.org/region/PE',
    }
    in_xls = 'Checklist_PEROU.xls'
else:
    # Just for the trip
    regions = {
        'J02_playa': 'https://ebird.org/hotspot/L3326274',
        'J02_Raf': 'https://ebird.org/hotspot/L2462199',
        'J03_Pomac' : 'https://ebird.org/hotspot/L11416923',
        'J03_Pomac2': 'https://ebird.org/hotspot/L14871716',
        'J04_Chaparri': 'https://ebird.org/hotspot/L604133',
        'J05_casupe' : 'https://ebird.org/hotspot/L54470165',
        'J05_abra_gav' : 'https://ebird.org/hotspot/L2341268',
        'J06_Rio_Chonta' : 'https://ebird.org/hotspot/L493415',
        'J06_Pu_Sangal' : 'https://ebird.org/hotspot/L7434187',
        'J06_Sangal' : 'https://ebird.org/hotspot/L4470950',
        'J07_Jesus' : 'https://ebird.org/hotspot/L4890252',
        'J07_Encañada' : 'https://ebird.org/hotspot/L21794436',
        'J07_Cruz_Conga' : 'https://ebird.org/hotspot/L493411',
        'J08_Endemic': 'https://ebird.org/hotspot/L4342331',
        'J08_Balsa': 'https://ebird.org/hotspot/L962441',
        'J08_Parrotlet': 'https://ebird.org/hotspot/L25652341',
        'J08_Barro_Negro': 'https://ebird.org/hotspot/L983325',
        'J09_Atuen': 'https://ebird.org/hotspot/L961859',
        'J09_Condor': 'https://ebird.org/hotspot/L983793',
        'J10_Kuelap': 'https://ebird.org/hotspot/L1137425',
        'J10_Utcu': 'https://ebird.org/hotspot/L983783',
        'J10_Espatula': 'https://ebird.org/hotspot/L20364615',
        'J10_Huembo': 'https://ebird.org/hotspot/L20357663',
        'J11_Pomacochas': 'https://ebird.org/hotspot/L9611856',
        'J11_Owlet': 'https://ebird.org/hotspot/L1849237',
        'J12_Alto_Nieva': 'https://ebird.org/hotspot/L31735009',
        'J12_Alto_Nieva2': 'https://ebird.org/hotspot/L1849274',
        'J12_Alto_Nieva3': 'https://ebird.org/hotspot/L1836547',
        'J12_Alto_Nieva4': 'https://ebird.org/hotspot/L4070147',
        'J13_Llanteria': 'https://ebird.org/hotspot/L1234162',
        'J13_Arena_Blanca': 'https://ebird.org/hotspot/L3637689',
        'J13_Humedales': 'https://ebird.org/hotspot/L2728917',
        'J14_Waqanki': 'https://ebird.org/hotspot/L1849133',
        'J14_Waq2': 'https://ebird.org/hotspot/L1652750',
        'J14_Quiscarr': 'https://ebird.org/hotspot/L976326',
        'J15_Aconabikh': 'https://ebird.org/hotspot/L2240938',
        'J15_Tunel': 'https://ebird.org/hotspot/L2273747',
        'J15_Yuri': 'https://ebird.org/hotspot/L2665114',
        'J15_Araca': 'https://ebird.org/hotspot/L3215723',
        'J16_Mirador': 'https://ebird.org/hotspot/L976324',
        'J16_Upaqui': 'https://ebird.org/hotspot/L976323',
        'J16_Plata': 'https://ebird.org/hotspot/L2455165',
    }

    in_xls = 'Checklist_NORD_PEROU.xls'


def extract_observation(url):
    if bird_name_in_english:
        website = requests.get(url + '/bird-list')
    else:
        website = requests.get(url + '/bird-list',
                            cookies={"I18N_LANGUAGE": "fr"})
    soup = BeautifulSoup(website.text, 'html.parser')

    result = {}

    for li in soup.find_all('li', class_='BirdList-list-list-item'):
        time_tag = li.find('time')
        if time_tag and time_tag.has_attr('datetime'):
            is_exotic = li.find('svg', class_='Icon--exoticEscapee') is not None
            if is_exotic:
                continue
            bird_name = li.find('span', class_='Species-common').get_text(strip=True)
            if 'sp.' in bird_name or '/' in bird_name or ' ou ' in bird_name or 'Hybride' in bird_name:
                continue
            result[bird_name] = dict(
                last_seen=datetime.strptime(time_tag['datetime'], '%Y-%m-%d %H:%M'))

    return result


infos = {}
for code, url in regions.items():
    print(f'Fetching {code}...')
    infos[code] = extract_observation(url)
    print(f'  -> {len(infos[code])} species found')


lookup = {}
for code, species_dict in infos.items():
    for name in species_dict:
        lookup.setdefault(normalize(name), set()).add(code)

rb = xlrd.open_workbook(in_xls, formatting_info=True)
sheet_index = rb.sheet_names().index('CHECKLIST')
rs = rb.sheet_by_index(sheet_index)

wb = xl_copy(rb)          # writable copy, preserves existing formatting
ws = wb.get_sheet(sheet_index)

row = 0
xls_names = set()
while row < rs.nrows:
    if bird_name_in_english:
        bird_name = rs.cell_value(row, 1)   # column B (index 1) = bird name
    else:
        bird_name = rs.cell_value(row, 3)   # column D (index 1) = bird name

    if bird_name == '':
        row += 1
        continue
    if bird_name != '':
        xls_names.add(normalize(str(bird_name)))
    codes = lookup.get(normalize(str(bird_name)))
    if codes:
        ws.write(row, 2, ', '.join(sorted(codes)))  # row index 2 = row "3"
    elif row > 1:
        print('ERROR FOR ', bird_name)

    row += 1
    if row > 2900:
       break

wb.save(out_xls)

# --- Find eBird species with no match in the XLS ---
unmatched_ebird = {name: codes for name, codes in lookup.items() if name not in xls_names}

print(f'\n{len(unmatched_ebird)} eBird species not found in the XLS:')
for name, codes in sorted(unmatched_ebird.items()):
    print(f' - {name}  (seen in: {", ".join(sorted(codes))})')
    