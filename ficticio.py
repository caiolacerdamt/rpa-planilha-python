import pandas as pd
import random
from faker import Faker
import numpy as np


fake = Faker()

data = {'Matrícula': [],
        'Magistrado': [],
        'Mês': [],
        'Audiências Realizadas': [],
        'Audiências realizadas no CEJUSC': [],
        'Despachos': [],
        'Despachos no CEJUSC': [],
        'Decisões': [],
        'Decisões no CEJUSC': [],
        'Julgamentos com Mérito': [],
        'Julgamentos com mérito no CEJUSC': [],
        'Julgamentos sem Mérito': [],
        'Julgamentos sem mérito no CEJUSC': [],
        'Excesso de Prazo Sentença': [],
        }

for _ in range(50):
    data['Matrícula'].append(fake.random_number(digits=8))
    data['Magistrado'].append(fake.name())
    data['Mês'].append(fake.date_time_between(start_date='-2y', end_date='now').strftime('%b-%y').lower())
    data['Audiências Realizadas'].append(random.randint(0, 999))
    data['Audiências realizadas no CEJUSC'].append(random.randint(0, 999))
    data['Despachos'].append(random.randint(0, 999))
    data['Despachos no CEJUSC'].append(random.randint(0, 999))
    data['Decisões'].append(random.randint(0, 999))
    data['Decisões no CEJUSC'].append(random.randint(0, 999))
    data['Julgamentos com Mérito'].append(random.randint(0, 999))
    data['Julgamentos com mérito no CEJUSC'].append(random.randint(0, 999))
    data['Julgamentos sem Mérito'].append(random.randint(0, 999))
    data['Julgamentos sem mérito no CEJUSC'].append(random.randint(0, 999))
    data['Excesso de Prazo Sentença'].append(random.randint(0,300))


df = pd.DataFrame(data)

df["Mês"] = pd.to_datetime(df["Mês"], format='%b-%y', errors='coerce').dt.strftime('%b-%y')
df = df.sort_values('Mês')

df['Matrícula'] = df['Matrícula'].astype(str)

total = df.select_dtypes(include=np.number).sum()
total['Magistrado'] = 'Total'

df = df._append(total, ignore_index=True)

df.to_excel('dados_ficticios.xlsx', index=False)
