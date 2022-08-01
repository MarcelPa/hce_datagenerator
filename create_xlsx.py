import copy
import csv
import datetime
import random
import string
from typing import Any, List, Dict

from faker import Faker
from faker_bloodpanel import BloodPanelProvider
import openpyxl
from openpyxl.styles import PatternFill

# settings for fake data
small_sheet_records = 10
large_sheet_records = 5000

# initialise some variables needed for both generatd sheets
today = datetime.date.today()
samples = ['bloed', 'uitstrijkje']
keys = ['naam', 'patient_id', 'datum', 'geboortedatum', 'straat', 'stad', 'type', 'genomen', 'ingang', 'Sodium', 'Potassium', 'Chloride', 'Bicarbonate', 'Urea', 'Magnesium', 'Total calcium', 'Hemoglobin', 'Covid-PCR']
column_widths = [18] + [14] * 3 + [18] * 2 + [14] * (len(keys) - 6) # header widths: name + others + street, city + rest

def generate_record(samples: List[str], fake: Faker) -> Dict[str, Any]:
    ''' Generate a record containing a random sample using Faker '''
    # map the sample to a generator function
    sample_random_method = {
        'bloed': lambda: fake.simple_blood_panel(add_units = True),
        'uitstrijkje': lambda: fake.covid_test(),
    }

    # start and populate a new record
    record = {}
    sample_type = random.choice(samples)
    record['type'] = sample_type
    record['genomen'], record['ingang'] = sorted((fake.time(), fake.time()))
    record = record | sample_random_method[sample_type]()

    # instead of True / False, we would like to have 'positief' / 'negatief'
    if sample_type == 'uitstrijkje':
        record['Covid-PCR'] = 'positief' if record['Covid-PCR'] else 'negatief'

    return record

def to_excel(sheet: List[Dict[str, Any]], keys: List[str], column_widths: List[int]) -> openpyxl.Workbook:
    # create a new excel workbook
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Labor'

    # add the header row and set the column widths
    for pos, key in enumerate(keys):
        ws[f'{string.ascii_uppercase[pos]}1'] = key.capitalize()
        ws.column_dimensions[f'{string.ascii_uppercase[pos]}'].width = column_widths[pos]

    # add the data rows
    for row, record in enumerate(small_sheet):
        for key in record.keys():
            pos = keys.index(key)
            ws[f'{string.ascii_uppercase[pos]}{row+2}'] = record[key]

    return wb

# initialise a new faker instance and add the bloodpanel provider
fake = Faker('nl_BE')
fake.add_provider(BloodPanelProvider)

# generate a list of patients including profiles
patients = []
for i in range(int(large_sheet_records * .8)):
    profile = fake.simple_profile()
    address = profile['address'].split('\n')
    patients.append({
        'naam': profile['name'],
        'patient_id': f'{(i+1) * 41232 % 100003:05d}',
        'datum': (today - datetime.timedelta(days=random.randint(14, 28))).isoformat(),
        'geboortedatum': profile['birthdate'].isoformat(),
        'straat': address[0],
        'stad': ' '.join(address[1:]),
    })

# generate the small sheet: pick a few patient at random and generate samples for them
small_sheet = []
small_i = int(small_sheet_records * .8) # 80% unique patients, 20% duplicates
for patient in patients[:small_i] + random.sample(patients[:small_i], k = int(small_sheet_records * .2)):
    record = copy.deepcopy(patient)
    record = record | generate_record(samples, fake)
    small_sheet.append(record)

# transform to excel and save it to a file
wb = to_excel(small_sheet, keys, column_widths)
wb.save('small_file.xlsx')

# generate the large sheet: pick all patients and generate samples for them
large_sheet = small_sheet # we are just adding new records to the small sheet
for patient in patients + random.sample(patients, k = int(large_sheet_records * .2) - small_sheet_records):
    record = copy.deepcopy(patient)
    record = record | generate_record(samples, fake)
    large_sheet.append(record)

# introduce some noise to the large sheet
for i in random.choices(list(range(small_sheet_records, large_sheet_records)), k = int(large_sheet_records * .05)):
    # define a list of 'noisable' keys
    noise_key = random.choice(['geboortedatum', 'type', 'Covid-PCR'])
    noise_map = {
        'geboortedatum': lambda cur: (datetime.date.fromisoformat(cur) - datetime.timedelta(days=random.randint(365 * 75, 365 * 150))).isoformat(),
        'type': lambda cur: 'bloed' if cur == 'uitstrijkje' else 'uitstrijkje',
        'Covid-PCR': lambda cur: random.choice(['pos', 'post', 'y', 'prositif', 'neg', '-', 'n', 'non'])
    }
    large_sheet[i][noise_key] = noise_map[noise_key](large_sheet[i][noise_key] if noise_key in large_sheet[i].keys() else None)

# transform to excel and save it to a file
wb = to_excel(large_sheet, keys, column_widths)
wb.save('large_file.xlsx')
