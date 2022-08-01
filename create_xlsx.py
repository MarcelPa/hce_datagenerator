import copy
import csv
import datetime
import random
import string

from faker import Faker
from faker_bloodpanel import BloodPanelProvider
import openpyxl
from openpyxl.styles import PatternFill

small_sheet_records = 10
large_sheet_records = 1000
today = datetime.date.today()

fake = Faker('nl_BE')
fake.add_provider(BloodPanelProvider)
samples = ['bloed', 'uitstrijkje']
sample_random_method = {
    'bloed': lambda: fake.simple_blood_panel(add_units = True),
    'uitstrijkje': lambda: fake.covid_test(),
}

patients = []
for i in range(int(small_sheet_records * .8)):
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

small_sheet = []
keys = ['naam', 'patient_id', 'datum', 'geboortedatum', 'straat', 'stad', 'type', 'genomen', 'ingang', 'Sodium', 'Potassium', 'Chloride', 'Bicarbonate', 'Urea', 'Magnesium', 'Total calcium', 'Hemoglobin', 'Covid-PCR']

for patient in patients + random.sample(patients, k = int(small_sheet_records * .2)):
    record = copy.deepcopy(patient)
    sample_type = random.choice(samples)
    record['type'] = sample_type
    record['genomen'], record['ingang'] = sorted((fake.time(), fake.time()))
    record = record | sample_random_method[sample_type]()
    small_sheet.append(record)

with open('small_sheet.csv', 'w') as outcsv:
    writer = csv.DictWriter(outcsv, fieldnames=keys)
    writer.writeheader()
    writer.writerows(small_sheet)

wb = openpyxl.Workbook()
ws = wb.active
ws.title = 'Labor'

# header widths: name + others + street, city + rest
column_widths = [18] + [14] * 3 + [18] * 2 + [14] * (len(keys) - 6)
for pos, key in enumerate(keys):
    ws[f'{string.ascii_uppercase[pos]}1'] = key.capitalize()
    ws.column_dimensions[f'{string.ascii_uppercase[pos]}'].width = column_widths[pos]

for row, record in enumerate(small_sheet):
    for key in record.keys():
        pos = keys.index(key)
        if pos == 17:
            ws[f'{string.ascii_uppercase[pos]}{row+2}'] = 'positief' if record[key] else 'negatief'
        else:
            ws[f'{string.ascii_uppercase[pos]}{row+2}'] = record[key]

wb.save('small.xlsx')

large_sheet = []
for i in range(int(large_sheet_records * .8) - int(small_sheet_records * .8)):
    profile = fake.simple_profile()
    address = profile['address'].split('\n')
    patients.append({
        'naam': profile['name'],
        'patient_id': f'{(i+1+small_sheet_records) * 41232 % 100003:05d}',
        'datum': (today - datetime.timedelta(days=random.randint(14, 28))).isoformat(),
        'geboortedatum': profile['birthdate'].isoformat(),
        'straat': address[0],
        'stad': ' '.join(address[1:]),
    })
for patient in patients + random.sample(patients, k = int(large_sheet_records * .2)):
    record = copy.deepcopy(patient)
    sample_type = random.choice(samples)
    record['type'] = sample_type
    record['genomen'], record['ingang'] = sorted((fake.time(), fake.time()))
    record = record | sample_random_method[sample_type]()
    large_sheet.append(record)

wb = openpyxl.Workbook()
ws = wb.active
ws.title = 'Labor'

# header widths: name + others + street, city + rest
column_widths = [18] + [14] * 3 + [18] * 2 + [14] * (len(keys) - 6)
for pos, key in enumerate(keys):
    ws[f'{string.ascii_uppercase[pos]}1'] = key.capitalize()
    ws.column_dimensions[f'{string.ascii_uppercase[pos]}'].width = column_widths[pos]

for row, record in enumerate(large_sheet):
    for key in record.keys():
        pos = keys.index(key)
        if pos == 17:
            ws[f'{string.ascii_uppercase[pos]}{row+2}'] = 'positief' if record[key] else 'negatief'
        else:
            ws[f'{string.ascii_uppercase[pos]}{row+2}'] = record[key]

wb.save('large.xlsx')
