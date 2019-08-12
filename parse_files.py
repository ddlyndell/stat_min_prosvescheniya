import pandas as pd
import os

# files directories
source_2018 = 'C:/minobr/source_2018/' 
source_2017 = 'C:/minobr/source_2017/'

files_2018 = os.listdir(source_2018)
files_2017 = os.listdir(source_2017)

# get all cases of parental rights in region
def parse_parental_rights(file, root, year):
    region = pd.read_excel(root + file,'Титульный лист', header=None).at[28,23]
    df = pd.read_excel(root + file, 'Раздел 5', skiprows=20, header=None).dropna(axis = 1, how = 'all')
    df.columns = ['factor', 'row_code', 'counter']
    df = df[df.row_code.isin([7,8,9,10,11,12,13])]
    df['region'] = region
    df['year'] = year
    return df

# get all cases of child abuse in region
def parse_child_abuse(file, root, year):
    region = pd.read_excel(root + file,'Титульный лист', header=None).at[28,23]
    df = pd.read_excel(root + file, 'Раздел 5', skiprows=20, header=None).dropna(axis = 1, how = 'all')
    df.columns = ['factor', 'row_code', 'counter']
    df = df[df.row_code.isin([38, 39, 40, 41, 42, 43])]
    df['region'] = region
    df['year'] = year
    return df
    
frames = []
for file in files_2018:
    frames.append(parse_parental_rights(file, source_2018, 2018))
total_2018 = pd.concat(frames)
total_2018.to_csv('Родители, лишенные прав 2018.csv', index = False)

frames = []
for file in files_2017:
    frames.append(parse_parental_rights(file, source_2017, 2017))
total_2017 = pd.concat(frames)
total_2017.to_csv('Родители, лишенные прав 2017.csv', index = False)

frames = []
for file in files_2017:
    frames.append(parse_child_abuse(file, source_2017, 2017))
total_ca_2017 = pd.concat(frames)
total_ca_2017.to_csv('Жестокое обращение с детьми 2017.csv', index = False)

frames = []
for file in files_2018:
    frames.append(parse_child_abuse(file, source_2018, 2018))
total_ca_2018 = pd.concat(frames)
total_ca_2018.to_csv('Жестокое обращение с детьми 2018.csv', index = False)
