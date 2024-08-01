# %% [markdown]
# 

# %% [markdown]
# #### Библиотеки для работы с данными
# 
# будет выдавать ошибки мол библиотека не найдена - удалите # со строки с соответствующей библиотекой

# %%
import sys
import re
#!{sys.executable} -m pip install recordlinkage -q
import recordlinkage as rl
#!{sys.executable} -m pip install xlwings -q
import xlwings as xw
#!{sys.executable} -m pip install openpyxl -q
import openpyxl
import numpy as np
import pandas as pd
from datetime import datetime, date
#!{sys.executable} -m pip install pyinstaller -q
import warnings
warnings.filterwarnings("ignore")

# %%
#sys.path

# %% [markdown]
# #### Чтение предоставленных к обработке файлов
# перез запуском - сверьте название эксель и листа (sheet_name)

# %% [markdown]
# этот юлок надо раскомментировать - пока только он рабоьает в гугл коллабе

# %%
#df = pd.read_excel('/content/Актуальный_отчет_292.xlsx', sheet_name = 'Лист1',  engine='openpyxl')
#dfFCCP = pd.read_excel('/content/Реестр ФССП.xlsx', engine='openpyxl')

# %% [markdown]
# этот, соответственно, нужно удалить ( на всякий случай проверьте - он изменен и может заработать, \
# скорее всего потребуется заменить название файла в нижних двух строках на путь (пкм по загруженному файлу - скопировать путь)

# %%
#в этом блоке функцией просматриваются эксель и берётся из него все записи с указанного листа и записываются в рабочий датафрейм
#Если надо просматривать более одного листа - в sheet_name=['первый лист','второй лист' и тд]
def xw_excel_to_df(file_path, sheet_name=None, sheet_range=None):
    app = xw.App(visible=False)
    book = app.books.open(file_path, read_only=True)

    if sheet_name is None:
        sheet_name = book.sheets[0].name

    if sheet_range is None:
        data = book.sheets[sheet_name]["A1"].expand().value
    else:
        data = book.sheets[sheet_name][sheet_range].value

    df = pd.DataFrame(data=data[1:], columns=data[0])

    book.close()
    app.quit()

    return df

df = xw_excel_to_df('Актуальный_отчет_292.xlsx')
dfFCCP = xw_excel_to_df('Реестр ФССП.xlsx')

# %%
#Просто вывод количества записей
print('Сравнение отчета 292 с реестром ФССП')
print('')
print('Изначальные данные в эксель-таблицах')
print('Количество строк в отчете 292:',len(df))
print('Количество записей (всего) в реестре:',len(dfFCCP))
print('-----------------------------------------------------')
print('')

# %% [markdown]
# фильтр для сбора отчета и реестра в 1 файл

# %%
def get_matches(row, df_lev_0, df_lev_1):
    if row['level_0'] in df_lev_0.index and row['level_1'] in df_lev_1.index:
        row292 = df_lev_0.loc[row['level_0']].reset_index(drop=True)
        rowFCCP = df_lev_1.loc[row['level_1']].reset_index(drop=True)
        return pd.concat([row292, rowFCCP], ignore_index=True)

    else:
        return pd.Series()

# %% [markdown]
# Функция для  описания разницы значений ИД\ИП

# %%
def process_data(df):
    for index, row in df.iterrows():
        if row['ИД формат 292'] != row['ИД формат ФССП']:
            df.at[index, 'Аналитика (python)'] = f"разное ИД у 1 ИП; "

        if row['ИП формат 292'] != row['ИП формат ФССП']:
            df.at[index, 'Аналитика (python)'] = f"разное ИП у 1 ИД; "


# %% [markdown]
# #### Предобработка данных
# 
# перед запуском проверьте корректность названия столбцов\
# Дата подачи в суд навсегда переименована в Дата подачи в ИО

# %%
# Замена двойных пробелов на одинарные
df.replace({'  ': ' '}, regex=True, inplace=True)
dfFCCP.replace({'  ': ' '}, regex=True, inplace=True)

#Переименования (если есть подобные дате подачи в ИО, тогда мне напишите, скажу шаги как 1 раз записать и каждый раз не править)
df.rename(columns={'Дата возбуждения': 'Дата возбуждения 292'}, inplace=True)
if 'Дата подачи в суд' in df.columns:
    df.rename(columns={'Дата подачи в суд': 'Дата подачи в ИО'}, inplace=True)
df.rename(columns={'Поставщик (взыскатель)': 'Взыскатель 292'}, inplace=True)

# %%
print('Предобработка данных')
print('В процессе предобработке данные очищены от двойных пробелов')
print('В отчете - "Дата возбуждения" переименована в "Дата возбуждения 292",')
print('"Дата подачи в суд" в "Дата подачи в ИО", а "Поставщик (взыскатель)" в "Взыскатель 292"')
print('В случае перманентного изменения дополнительных столбцов - сообщите мне')
print('-------------------------------------------------------------------------------------------')

# %% [markdown]
# Даты

# %%
# Все столбцы с датами в отчете
date_columns_292 = ['Дата вынесения решения', 'Дата вступления в силу', 'Дата формирования ИП', 'Дата подачи в ИО',
                     'Период с', 'Период по', 'Дата получения листа на руки', 'Дата рождения ответчика']

# Все столбцы с датами в реестре
date_columns_FCCP = ['Дата возбуждения', 'Дата выдачи ИД']

print('')
print('Была произведена обработка дат в следующих столбцах:')
print('Для Отчета:')
print(' "Дата вынесения решения", "Дата вступления в силу", "Дата формирования ИП", "Дата подачи в ИО",')
print('"Период с", "Период по", "Дата получения листа на руки", "Дата рождения ответчика"')
print('Для реестра:')
print(' "Дата возбуждения", "Дата выдачи ИД"')
print('')

# %%
df[date_columns_292].replace({' ': ''}, regex=True, inplace=True)

for column in date_columns_292:
    df[column] = df[column].fillna('2050-12-30')

# Преобразование всех значений в 'datetime64[ns]', замена недопустимых значений на '2050-12-30'
for column in date_columns_292:
    df[column] = pd.to_datetime(df[column], errors='coerce')

# Преобразование дат, чьи года менее 1900 и выше 2050, в '2050-12-30'
for column in date_columns_292:
    df[column] = df[column].apply(lambda x: '2050-12-30' if pd.to_datetime(x, errors='coerce').year < 1900 or pd.to_datetime(x, errors='coerce').year > 2050 else x)

# Вывод всех строк, не соответствующих формату r'(\d{4}-\d{2}-\d{2})'
#for column in date_columns_292:
#    incorrect_dates = df[~df[column].astype(str).str.match(r'^\d{4}-\d{2}-\d{2}$')]
#    print(f"Строки с некорректным форматом даты в столбце {column}:")
#    display(incorrect_dates)

# %%
print('А именно:')
print('- Чистка от пробелов в столбцах')
print('- Замена пустых ячеек на "2050-12-30"')
print('- Замена годов ниже 1900 и выше 2050 на 2050')
print('- Преобразование в формат даты (из строчного и формата дата + время)')
print(' В остальных столбцах произведена замена пустых значений на "[EMPTY]"')
print('')

# %%
dfFCCP[date_columns_FCCP].replace({' ': ''}, regex=True, inplace=True)

for column in date_columns_FCCP:
    dfFCCP[column] = dfFCCP[column].fillna('2050-12-30')

for column in date_columns_FCCP:
    dfFCCP[column] = pd.to_datetime(dfFCCP[column], errors='coerce').fillna('2050-12-30')

for column in date_columns_FCCP:
    dfFCCP[column] = dfFCCP[column].apply(lambda x: '2050-12-30' if pd.to_datetime(x, errors='coerce').year < 1900 or pd.to_datetime(x, errors='coerce').year > 2050 else x)

#for column in date_columns_FCCP:
#    incorrect_dates = dfFCCP[~dfFCCP[column].astype(str).str.match(r'^\d{4}-\d{2}-\d{2}$')]
#    print(f"Строки с некорректным форматом даты в столбце {column}:")
#    display(incorrect_dates)

# %%
# Проверка и замена непустых значений
df.replace(r'^\s*$', '[EMPTY]', regex=True, inplace=True)
dfFCCP.replace(r'^\s*$', '[EMPTY]', regex=True, inplace=True)

# %% [markdown]
# Должник\взыскатель

# %%
#Для 292
df['Взыскатель 292'] = df['Взыскатель 292'].str.lower().str.strip()
df['Взыскатель 292'] = df['Взыскатель 292'].fillna('ПАО "Т Плюс"')
df['Взыскатель 292'] = df['Взыскатель 292'].apply(lambda x: 'Оренбургский филиал АО "ЭнергосбыТ Плюс"' if 'энергосбыт плюс' in x else ('ПАО "Т Плюс"' if 'т плюс' in x else x))

# Преобразование многообразия взыскателей в Оренбургский филиал АО "ЭнергосбыТ Плюс" или ПАО "Т Плюс"
dfFCCP['Взыскатель ФССП'] = dfFCCP['Взыскатель'].str.lower().str.strip()
dfFCCP['Взыскатель ФССП'] = dfFCCP['Взыскатель ФССП'].apply(lambda x: 'Оренбургский филиал АО "ЭнергосбыТ Плюс"' if 'энергосбыт плюс' in x else 'ПАО "Т Плюс"')



df['Должник 292'] = np.where(df['Ответчик'].notnull(), df['Ответчик'].str.lower(), df['ФИО должника'].str.lower())
df['Должник 292']=df['Должник 292'].str.strip()

dfFCCP['Должник ФССП'] = np.where(dfFCCP['Должник'].notnull(), dfFCCP['Должник'].str.lower(),
                               dfFCCP['Фамилия должника'].str.lower() + ' ' + dfFCCP['Имя должника'].str.lower() + ' ' + dfFCCP['Отчество должника'].str.lower())
dfFCCP['Должник ФССП']=dfFCCP['Должник ФССП'].str.strip()

# %%
print('Поставщик в отчете был либо изменен на Оренбургский филиал АО "ЭнергосбыТ Плюс", либо ПАО "Т Плюс",')
print('либо оставленн без изменений для всяких ООО и ТСЖ')
print('В реестре всё, что не оренбургский филиал - всё ПАО')
print('')
print('Должник в отчете почищен от пробелов')
print('В реестре либо взят из столбца "Должник", либо собран из ФИО.')
print('')

# %% [markdown]
# № ИП\судебного дела (292)

# %%
df.replace('//', '/', regex=True, inplace=True)
df['Номер ИП'].replace(' ', '', regex=True, inplace=True)
df['Номер ИП'] = df['Номер ИП'].str.strip()
# Приведение ИП к единому формату
pattern_ip = r'(\d{1,6}/\d{2,4}/\d{5,6})(-ИП)?'
df['ИП формат 292'] = df['Номер ИП'].str.extract(pattern_ip, expand=False).apply(lambda x: f"{x[0]}-ИП" if pd.notnull(x[0]) else None, axis=1)

df['Номер судебного дела'] = df['Номер судебного дела'].str.replace(' ', '')
df['Номер судебного дела'] = df['Номер судебного дела'].str.strip()

# Приедение ИД к единому формату
pattern_id = r'(\d{1,2}-\d{3,4}/\d{2,3})(/\d{2,4})|\
                (\d{1,2}-\d{1,2}-\d{4,5})(/\d{2,4})|\
                (\D+\d{9})|\
                (А47-\d{3,4}/\d{2,4})|\
                (\d{1,3}-\d{1,5})(?:/\d{2,4})?|\
                (\d{1,2}-{2,5}//d{2,4})|\
                (А47-\d{3,4}/\d{2,4})|\
                (\d{1,2}-{2,5}//d{1,3}//d{2,4})|\
                (\D+№\d{9})|\
                (\d{1,3}-\d{1,5})(?:/\d{2,4})?'
                
df['ИД формат 292'] = df['Номер судебного дела'].str.extract(pattern_id, expand=False).apply(lambda x: x[0] if pd.notnull(x[0]) else '[EMPTY]', axis=1)

# %%
print('В ИП почищены пробелы, // и видоизменено так, чтобы значение всегда имело "-ИП" на конце')
print('Для ИД всё, что не входит в паттерн ниже выходит в ошибки/пустоты')
print('pattern_id = r(\d{1,2}-\d{3,4}/\d{2,3})(/\d{2,4})|')
print('                (\d{1,2}-\d{1,2}-\d{4,5})(/\d{2,4})|')
print('                (\D+\d{9})|')
print('                (А47-\d{3,4}/\d{2,4})|')
print('                (\d{1,3}-\d{1,5})(?:/\d{2,4})?|')
print('                (\d{1,2}-{2,5}//d{2,4})|')
print('                (А47-\d{3,4}/\d{2,4})|')
print('                (\d{1,2}-{2,5}//d{1,3}//d{2,4})|')
print('                (\D+№\d{9})|')
print('                (\d{1,3}-\d{1,5})(?:/\d{2,4})?')
print('Для уточнения \d{...} - возможное количество цифр подряд')
print('\D+ - любое количество букв')
print('? - значение в круглой скобке не обязано существовать')
print('')
print('----------------------------------------------------------------------------------------------')
print('')

# %% [markdown]
# № ИП\судебного дела (FCCP)

# %%
dfFCCP.replace('//', '/', regex=True, inplace=True)
dfFCCP['Регистрационный номер ИП'].replace(' ', '', regex=True, inplace=True)
dfFCCP['Регистрационный номер ИП']=dfFCCP['Регистрационный номер ИП'].str.strip()
dfFCCP['ИП формат ФССП'] = dfFCCP['Регистрационный номер ИП'].str.extract(pattern_ip, expand=False).apply(lambda x: f"{x[0]}-ИП" if pd.notnull(x[0]) else None, axis=1)

dfFCCP['Номер ИД'] = dfFCCP['Номер ИД'].str.replace(' ', '')
dfFCCP['Номер ИД'] = dfFCCP['Номер ИД'].str.strip()
dfFCCP['ИД формат ФССП'] = dfFCCP['Номер ИД'].str.extract(pattern_id, expand=False).apply(lambda x: x[0] if pd.notnull(x[0]) else '[EMPTY]', axis=1)

# Фильтрация по статусу Возбуждено или Подано
filtered_df292_1 = df[df["Статус ИП"].isin(["Возбуждено", "Подано"])].reset_index()

# %%
# Проверка и замена пустых значений
filtered_df292_1.fillna('[EMPTY]', inplace=True)
dfFCCP.fillna('[EMPTY]', inplace=True)

# %%
print('Количество строк после первоначальной фильтрации')
print('Количесво строк в отфильтрованном (только со статусом возбуждено/подано) отчете 292:',len(filtered_df292_1))
print('Количество строк в отфильтрованном реестре:',len(dfFCCP))
print('-----------------------------------------------------')
print('')

# %% [markdown]
# #### Данные без ИД и ИП
# 
# Здесь я отрезаю данные, у которых нет ни ИП, ни ИД, а так же строки без должников - для дальнейших уточнений

# %%
# отбор данных без взыскателя в отчете
empty_data_292_debtor = filtered_df292_1[(filtered_df292_1['Должник 292'].isnull())|(filtered_df292_1['Должник 292'] == '[EMPTY]')]

# Убираем их из основного поиска
filtered_df292_02 = filtered_df292_1[~filtered_df292_1.index.isin(empty_data_292_debtor.index)]

# отбор данных без взыскателя в реестре
empty_data_FCCP_debtor = dfFCCP[(dfFCCP['Должник ФССП'].isnull())|(dfFCCP['Должник ФССП'] == '[EMPTY]')]

# Убираем их из основного поиска
dfFCCP_02 = dfFCCP[~dfFCCP.index.isin(empty_data_FCCP_debtor.index)]


# %%
# отбор данных без ИД и ИП в отчете
empty_data_292 = filtered_df292_1[((filtered_df292_1['ИД формат 292'].isnull())| (filtered_df292_1['ИД формат 292'] == '[EMPTY]')) & ((filtered_df292_1['ИП формат 292'].isnull())| (filtered_df292_1['ИП формат 292'] == '[EMPTY]'))]

# Убираем их из основного поиска
filtered_df292_2 = filtered_df292_02[~filtered_df292_02.index.isin(empty_data_292.index)]

# отбор данных без ИД и ИП в реестре
empty_data_FCCP = dfFCCP[((dfFCCP['ИД формат ФССП'].isnull())| (dfFCCP['ИД формат ФССП'] == '[EMPTY]')) & ((dfFCCP['ИП формат ФССП'].isnull())| (dfFCCP['ИП формат ФССП'] == '[EMPTY]'))]

# Убираем их из основного поиска
dfFCCP_2 = dfFCCP_02[~dfFCCP_02.index.isin(empty_data_FCCP.index)]

# %%
filtered_ed_matches_12 = empty_data_292.loc[
    ((empty_data_292['ИД формат 292'] != '[EMPTY]') | 
    (empty_data_292['ИП формат 292'] != '[EMPTY]'))
]

# Вывод выбранных данных
#display(filtered_ed_matches_12.loc[:, ['Дата подачи в ИО', 'Должник 292', 'ИД формат 292', 'ИП формат 292',  'Взыскатель 292']])

# %%
filtered_ed_matches_22 = empty_data_FCCP.loc[
    ((empty_data_FCCP['ИД формат ФССП'] != '[EMPTY]') |
    (empty_data_FCCP['ИП формат ФССП'] != '[EMPTY]'))]

# Вывод выбранных данных
#display(filtered_ed_matches_22.loc[:, ['Дата возбуждения', 'Должник ФССП', 'ИД формат ФССП', 'Взыскатель ФССП', 'ИП формат ФССП']])

# %%
print('Здесь я отрезаю данные, у которых нет ни ИП, ни ИД, а так же строки без должников - для дальнейших уточнений')
print('Количесво строк без ИД и ИП в отчете 292:',len(empty_data_292))
print('Количество строк без ИД и ИП в реестре:',len(empty_data_FCCP))
print('----------------------------------------------------')
print('Количесво строк в отфильтрованном отчете 292:',len(filtered_df292_2))
print('Количество строк в отфильтрованном реестре:',len(dfFCCP_2))
print('----------------------------------------------------')
print('')


# %% [markdown]
# сравнение с пустыми ИД и ИП

# %%
indexer = rl.Index()
indexer.block(left_on=['Должник 292', 'Взыскатель 292'], right_on=['Должник ФССП', 'Взыскатель ФССП'])
debtor = indexer.index(empty_data_292, empty_data_FCCP)

compare = rl.Compare()
compare.exact('Должник 292', 'Должник ФССП', label = 'Должник итог')
compare.exact('Взыскатель 292', 'Взыскатель ФССП', label = 'Взыскатель итог')
features = compare.compute(debtor, empty_data_292, empty_data_FCCP)
potential_matches_01 = features[features.sum(axis=1) > 1].reset_index()
print(len(debtor), len(potential_matches_01))

if potential_matches_01.empty:
    ed_matches = pd.DataFrame()
else:
    ed_matches = potential_matches_01.apply(get_matches, df_lev_0=empty_data_292, df_lev_1=empty_data_FCCP, axis=1).reset_index(drop=True)
    ed_matches.columns = empty_data_292.columns.append(empty_data_FCCP.columns)
    if 'index' in ed_matches.columns:
        ed_matches = ed_matches.drop('index', axis=1)

empty_data_292_1 = empty_data_292.copy().drop(potential_matches_01['level_0'].array).reset_index(drop=True)
empty_data_FCCP_1 = empty_data_FCCP.copy().drop(potential_matches_01['level_1'].array).reset_index(drop=True)
print('Количество найденных соответствий среди пустых:', potential_matches_01)
print('')

# %%
indexer = rl.Index()
indexer.block(left_on=['Должник 292', 'Взыскатель 292'], right_on=['Должник ФССП', 'Взыскатель ФССП'])
debtor = indexer.index(empty_data_292_1, dfFCCP_2)

compare = rl.Compare()
compare.exact('Должник 292', 'Должник ФССП', label = 'Должник итог')
compare.exact('Взыскатель 292', 'Взыскатель ФССП', label = 'Взыскатель итог')
features = compare.compute(debtor, empty_data_292_1, dfFCCP_2)
potential_matches_02 = features[features.sum(axis=1) == 2].reset_index()
print(len(debtor), len(potential_matches_02))

if potential_matches_02.empty:
    ed_matches_2 = pd.DataFrame()
else:
    ed_matches_2 = potential_matches_02.apply(get_matches, df_lev_0=empty_data_292_1, df_lev_1=dfFCCP_2, axis=1).reset_index(drop=True)
    ed_matches_2.columns = empty_data_292_1.columns.append(dfFCCP_2.columns)
    if 'index' in ed_matches_2.columns:
        ed_matches_2 = ed_matches_2.drop('index', axis=1)

empty_data_292_2 = empty_data_292_1.copy().drop(potential_matches_02['level_0'].array).reset_index(drop=True)
print('Количество найденных соответствий некорректных записей отчета с реестром:', len(potential_matches_02))
print('P.S. - найденные не будут вырезаться из вариантов в реестре')
print('')

# %%
filtered_ed_matches_2 = ed_matches_2.loc[
    ((ed_matches_2['ИД формат ФССП'] != '[EMPTY]') & 
    (ed_matches_2['ИП формат ФССП'] != '[EMPTY]')) | 
    ((ed_matches_2['ИД формат 292'] != '[EMPTY]') & 
    (ed_matches_2['ИП формат 292'] != '[EMPTY]'))
]

# Вывод выбранных данных
#display(filtered_ed_matches_2.loc[:, ['Дата подачи в ИО', 'Дата возбуждения','Должник 292', 'ИД формат 292', 'Взыскатель 292', 'ИП формат 292', 'Должник ФССП', 'ИД формат ФССП', 'Взыскатель ФССП', 'ИП формат ФССП']])

# %%
combined_ed_matches = pd.concat([ed_matches, ed_matches_2], ignore_index=True)
print('Всего найдено соответствий по должнику и взыскателю среди пустых - ', len(combined_ed_matches))
print('')
print('Что там дубликатов соразмерно разнообразию дат')
print('Я их чистить не стал т к соответствия он нормально нашел - а данные в БД в любом случае желательно корректно заполнить..)')
print('Надо будет ограничить поиск непосредственно по максимально большим датам - напишите')
print('Аналогично про сверку некорректных дат для некорректно заполненных сравниваемых строк')
print('')
print('---------------------------------------------------------------------------------------------------------------------------')
print('')

# %% [markdown]
# все пустые вместе с частичными соответствиями будут записаны в 1 файл, блок с этим в конце частичных соответствий

# %% [markdown]
# #### Блоки полного сравнения
# 
# перед запуском проверьте корректность названия столбцов

# %%
indexer = rl.Index()
indexer.block(left_on=['Должник 292', 'ИД формат 292', 'Взыскатель 292'], right_on=['Должник ФССП', 'ИД формат ФССП', 'Взыскатель ФССП'])
debtor = indexer.index(filtered_df292_2, dfFCCP_2)

compare = rl.Compare()
compare.string('Должник 292', 'Должник ФССП', label = 'Должник итог')
compare.string('ИД формат 292', 'ИД формат ФССП', label = 'Номер ИД итог')
compare.string('ИП формат 292', 'ИП формат ФССП', label = 'Номер ИП итог')
compare.string('Взыскатель 292', 'Взыскатель ФССП', label = 'Взыскатель итог')
features = compare.compute(debtor, filtered_df292_1, dfFCCP)
potential_matches = features[(features['Должник итог'] == 1) & (features['Номер ИД итог'] == 1) & (features['Взыскатель итог'] == 1)].reset_index()

print('Блок полного сравнения')
print('Количество найденных соответствий по ИД:', len(potential_matches))
print('Здесь выделяются строки, чьи должник ИД и взыскатель равны')

# %% [markdown]
# #### Немного анализа через питон
# 
# 1 блок - создает эксельку с найденными строками в отчетах (level_0 - отчет, level_1 - реестр)\
# 2 блок - просмотр рассматриваемых строк в отчете на определенной строке\
# 3 блок - просмотр рассматриваемых строк в реестре на определенной строке\
# если надо - раскомментируете

# %%
                        # создает эксель с найденными строками в отчетах (level_0 - отчет, level_1 - реестр)
#full_matches = potential_matches.apply(get_matches, df_lev_0=filtered_df292_2, df_lev_1=dfFCCP_2, axis=1).reset_index(drop=True)
#full_matches.columns = filtered_df292_2.columns.append(dfFCCP_2.columns)
#full_matches = full_matches.drop('index', axis=1)
#print('Уникальных ФИО:', full_matches['Должник 292'].nunique())

#with pd.ExcelWriter('Полные соответствия (строки).xlsx') as writer:
#    full_matches.to_excel(writer, sheet_name='Полные соответствия', index=False, engine='openpyxl', startrow=0, startcol=0)
#    potential_matches.to_excel(writer, sheet_name='строки', index=False, engine='openpyxl', startrow=0, startcol=0)

# %%
# print(filtered_df292_2.loc[136, ['Должник 292', 'ИД формат 292', 'Взыскатель 292', 'ИП формат 292']])

# %%
#print(dfFCCP_2.loc[71561, ['Должник ФССП', 'ИД формат ФССП', 'Взыскатель ФССП', 'ИП формат ФССП']])

# %%
non_matches_1_1 = filtered_df292_2.copy().drop(potential_matches['level_0'].array).reset_index(drop=True)
non_fc_matches_1_1 = dfFCCP_2.copy().drop(potential_matches['level_1'].array).reset_index(drop=True)

matches = potential_matches.apply(get_matches, df_lev_0=filtered_df292_2, df_lev_1=dfFCCP_2, axis=1)
matches.columns = filtered_df292_1.columns.append(dfFCCP.columns)
matches = matches.drop('index', axis=1)
process_data(matches)

# %%
indexer = rl.Index()
indexer.block(left_on=['Должник 292', 'ИП формат 292', 'Взыскатель 292'], right_on=['Должник ФССП', 'ИП формат ФССП', 'Взыскатель ФССП'])
debtor = indexer.index(non_matches_1_1, non_fc_matches_1_1)

compare = rl.Compare()
compare.string('Должник 292', 'Должник ФССП', label = 'Должник итог')
compare.string('ИД формат 292', 'ИД формат ФССП', label = 'Номер ИД итог')
compare.string('ИП формат 292', 'ИП формат ФССП', label = 'Номер ИП итог')
compare.string('Взыскатель 292', 'Взыскатель ФССП', label = 'Взыскатель итог')
features2 = compare.compute(debtor, non_matches_1_1, non_fc_matches_1_1)
potential_matches2 = features2[(features2['Должник итог'] == 1) & (features2['Номер ИП итог'] == 1) & (features2['Взыскатель итог'] == 1)].reset_index()

print('Количество найденных соответствий по ИП:', len(potential_matches2))
print('Здесь выделяются строки, чьи должник ИП и взыскатель равны')

# %% [markdown]
# Упрощенный аналог анализа для ИП
# 
# Если что, это можно использовать в любых блоках:
# - добавляете три блока через + слева наверху
# - копируете эти строки туда
# - удаляете все знаки # (т е знак комментария)
# - заменяете 'название датафрейма'.loc на названия из строки debtor = indexer.index(датафрейм отчета, датафрейм реестра)
# - номера строк берутся на основе level_0 и level_1 соответственно

# %%
# potential_matches2

# %%
# print(non_matches.loc[99, ['Должник 292', 'Взыскатель 292', 'ИД формат 292', 'ИП формат 292']])

# %%
# print(non_fc_matches.loc[229877, ['Должник ФССП', 'Взыскатель ФССП', 'ИД формат ФССП', 'ИП формат ФССП']])

# %%
non_mathes_1_2 = non_matches_1_1.drop(potential_matches2['level_0'].array)
non_mathes_1_2 = non_mathes_1_2.drop('index',axis=1)
non_fc_matches_1_2 = non_fc_matches_1_1.drop(potential_matches2['level_1'].array).reset_index(drop=True)

matches2 = potential_matches2.apply(get_matches, df_lev_0=non_matches_1_1, df_lev_1=non_fc_matches_1_1, axis=1).reset_index(drop=True)
matches2.columns = non_matches_1_1.columns.append(non_fc_matches_1_1.columns)
matches2 = matches2.drop('index', axis=1)
process_data(matches2)

# %%
# Сбор всех найденных строк воедино
matches_all = pd.concat([matches, matches2], ignore_index=True)

# Отрезаем некорректные даты и завершаем обработку найденных соответствий
matches_all_1 = matches_all.loc[(matches_all['Дата подачи в ИО'] < matches_all['Дата возбуждения'])]

# Подготовка к формированию неявных соответствий 1
non_matches_fin_0 = matches_all[~matches_all.index.isin(matches_all_1.index)]
non_matches_fin_0['Аналитика (python)'] = 'ошибка в дате'
non_matches_fin_0['Несоответствие'] = 'Дата возбуждения больше даты подачи в ИО'

print('Количество строк в сборной таблице после чистки некорректных дат:', len(matches_all_1))
print('Количество строк с некорректными датами:', len(non_matches_fin_0))
print('P.S. - заполнитель (2050 год) я больше не вырезаю - теперь он портит статистику со всем возможным усердием')
print('-----------------------------------------------------------------')
print('')

# %%
print('Итоговое количество записей отчет:', len(matches_all_1))
print('Итоговое количество должников отчет:', matches_all_1['ID ИП'].nunique())
print('Изменил проверку уникальных на проверку по ID ИП')
print('')
print('--------------------------------------------------------------------')
print('')

# %% [markdown]
# как вы видите - разница есть\
# часть этого приходится на разные заявления на одного должника, но часть из этого, возможно, дубли\
# или более точно - записи из отчета, которые нашли более одного соответствия в реестре на основе предоставленных 4 параметров\
# без доп. параметров отделить одно от другого можно только ручками\
# если у вас есть предложение, какой ещё столбец можно использовать для отрезания подобных записей - с удовольствием добью автоматизацию.

# %%
print('как вы видите - разница есть')
print('часть этого приходится на разные заявления на одного должника, но часть из этого, возможно, дубли')
print('или более точно - записи из отчета, которые нашли более одного соответствия в реестре на основе предоставленных 4 параметров')
print('без доп. параметров отделить одно от другого можно только ручками')
print('если у вас есть предложение, какой ещё столбец можно использовать для отрезания подобных записей - с удовольствием добью автоматизацию.')
print('----------------------------------------------------------------------------------------------------------------------------------------')
print('')

# %% [markdown]
# #### Блок частичного сравнения
# 
# перед запуском проверьте корректность названия столбцов

# %% [markdown]
# Несоответствие взыскателей (ИП)

# %%
indexer = rl.Index()
indexer.block(left_on=['Должник 292', 'ИП формат 292'], right_on=['Должник ФССП', 'ИП формат ФССП'])
debtor = indexer.index(non_mathes_1_2, non_fc_matches_1_2)

compare = rl.Compare()
compare.string('Должник 292', 'Должник ФССП', label = 'Должник итог')
compare.string('ИП формат 292', 'ИП формат ФССП', label = 'Номер ИП итог')
compare.string('Взыскатель 292', 'Взыскатель ФССП', label = 'Взыскатель итог')
features3 = compare.compute(debtor, non_mathes_1_2, non_fc_matches_1_2)
potential_matches3 = features3[(features3['Должник итог'] == 1) & (features3['Номер ИП итог'] == 1) & (features3.sum(axis=1) <3)].reset_index()

print('Блок частичного сравнения')
print('Количество найденных соответствий запросу (Несоответствие взыскателей (ИП)):', len(potential_matches3))
print('Сравнение идёт по полному совпадению должника и ИП')
print('')

# %%
non_mathes_1_3 = non_mathes_1_2.drop(potential_matches3['level_0'].array)
if 'index' in non_mathes_1_3.columns:
    non_mathes_1_3 = non_mathes_1_3.drop('index',axis=1)
non_fc_matches_1_3 = non_fc_matches_1_2.drop(potential_matches3['level_1'].array).reset_index(drop=True)

if len(potential_matches3)>0:
    non_matches_fin_1 = potential_matches3.apply(get_matches, df_lev_0=non_mathes_1_2, df_lev_1=non_fc_matches_1_2, axis=1).reset_index(drop=True)
    non_matches_fin_1.columns = non_mathes_1_2.columns.append(non_fc_matches_1_2.columns)
    if 'index' in non_matches_fin_1.columns:
        non_matches_fin_1 = non_matches_fin_1.drop('index', axis=1)
    process_data(non_matches_fin_1)
    non_matches_fin_1['Несоответствие'] = 'Взыскатель (по ИП)'
else:
    non_matches_fin_1 = pd.DataFrame()

# %% [markdown]
# Несоответствие взыскателей (физ лица)

# %%
indexer = rl.Index()
indexer.block(left_on=['Должник 292', 'ИД формат 292'], right_on=['Должник ФССП', 'ИД формат ФССП'])
debtor = indexer.index(non_mathes_1_3, non_fc_matches_1_3)

compare = rl.Compare()
compare.string('Должник 292', 'Должник ФССП', label = 'Должник итог')
compare.string('ИП формат 292', 'ИП формат ФССП', label = 'Номер ИП итог')
compare.string('ИД формат 292', 'ИД формат ФССП', label = 'Номер ИД итог')
compare.string('Взыскатель 292', 'Взыскатель ФССП', label = 'Взыскатель итог')
features4 = compare.compute(debtor, non_mathes_1_3, non_fc_matches_1_3)
potential_matches4 = features4[(features4['Должник итог'] == 1) & (features4['Номер ИД итог'] == 1)].reset_index()

print('Количество найденных соответствий запросу (Несоответствие взыскателей (ИД)):',len(potential_matches4))
print('Сравнение идёт по полному совпадению должника и ИД')
print('')

# %%
non_mathes_1_4 = non_mathes_1_3.drop(potential_matches4['level_0'].array)
if 'index' in non_mathes_1_4.columns:
    non_mathes_1_4 = non_mathes_1_4.drop('index',axis=1)
non_fc_matches_1_4 = non_fc_matches_1_3.drop(potential_matches4['level_1'].array).reset_index(drop=True)

if len(potential_matches4)>0:
    non_matches_fin_2 = potential_matches4.apply(get_matches, df_lev_0=non_mathes_1_3, df_lev_1=non_fc_matches_1_3, axis=1).reset_index(drop=True)
    non_matches_fin_2.columns = non_mathes_1_3.columns.append(non_fc_matches_1_3.columns)
    if 'index' in non_matches_fin_2.columns:
        non_matches_fin_2 = non_matches_fin_2.drop('index', axis=1)
    process_data(non_matches_fin_2)
    non_matches_fin_2['Несоответствие'] = 'Взыскатель (по ИД)'
else:
    non_matches_fin_2 = pd.DataFrame()

# %% [markdown]
# #### Совпадение по ИД > 90%

# %%
indexer = rl.Index()
indexer.block(left_on=['Должник 292', 'Взыскатель 292'], right_on=['Должник ФССП', 'Взыскатель ФССП'])
debtor = indexer.index(non_mathes_1_4, non_fc_matches_1_4)

compare = rl.Compare()
compare.string('Должник 292', 'Должник ФССП', label = 'Должник итог')
compare.string('ИП формат 292', 'ИП формат ФССП', label = 'Номер ИП итог')
compare.string('ИД формат 292', 'ИД формат ФССП', label = 'Номер ИД итог')
compare.string('Взыскатель 292', 'Взыскатель ФССП', label = 'Взыскатель итог')
features5 = compare.compute(debtor, non_mathes_1_4, non_fc_matches_1_4)
potential_matches5 = features5[(features5['Должник итог'] == 1) 
                               & (features5['Взыскатель итог'] == 1) 
                               & (features5['Номер ИД итог'] >= 0.8) 
                               & (features5['Номер ИП итог'] >= 0.5)].reset_index()

print('Количество найденных соответствий запросу (Совпадение по ИД > 80%):',len(potential_matches5))
print('Сравнение идёт по полному совпадению должника и взскателя, где схождение по ИД превышает 80%')
print('а ИП совпадает более чем на 50%')
print('Обычно при таком соотношении выпадают разного рода ошибки в буквах')
print('В нашем случае - поиск родственников по ИД')
print('')

# %%
non_mathes_1_5 = non_mathes_1_4.drop(potential_matches5['level_0'].array)
if 'index' in non_mathes_1_5.columns:
    non_mathes_1_5 = non_mathes_1_5.drop('index',axis=1)
non_fc_matches_1_5 = non_fc_matches_1_4.drop(potential_matches5['level_1'].array).reset_index(drop=True)

if len(potential_matches5)>0:
    non_matches_fin_3 = potential_matches5.apply(get_matches, df_lev_0=non_mathes_1_4, df_lev_1=non_fc_matches_1_4, axis=1).reset_index(drop=True)
    non_matches_fin_3.columns = non_mathes_1_4.columns.append(non_fc_matches_1_4.columns)
    if 'index' in non_matches_fin_3.columns:
        non_matches_fin_3 = non_matches_fin_3.drop('index', axis=1)
    process_data(non_matches_fin_3)
    non_matches_fin_3['Несоответствие'] = 'Совпадение по ИД > 80%'
else:
    non_matches_fin_2 = pd.DataFrame()

# %% [markdown]
# #### Совпадение по ИП > 90%

# %%
indexer = rl.Index()
indexer.block(left_on=['Должник 292', 'Взыскатель 292'], right_on=['Должник ФССП', 'Взыскатель ФССП'])
debtor = indexer.index(non_mathes_1_5, non_fc_matches_1_5)

compare = rl.Compare()
compare.string('Должник 292', 'Должник ФССП', label = 'Должник итог')
compare.string('ИП формат 292', 'ИП формат ФССП', label = 'Номер ИП итог')
compare.string('ИД формат 292', 'ИД формат ФССП', label = 'Номер ИД итог')
compare.string('Взыскатель 292', 'Взыскатель ФССП', label = 'Взыскатель итог')
features6 = compare.compute(debtor, non_mathes_1_5, non_fc_matches_1_5)
potential_matches6 = features6[(features6['Должник итог'] == 1) 
                               & (features6['Взыскатель итог'] == 1)
                               & (features6['Номер ИП итог'] >= 0.8)
                               & (features6['Номер ИД итог'] >= 0.5)].reset_index()

print('Количество найденных соответствий запросу (Совпадение по ИД > 80%):',len(potential_matches6))
print('Сравнение идёт по полному совпадению должника и взскателя, где схождение по ИП превышает 80%')
print('а ИД совпадает более чем на 50%')
print('Обычно при таком соотношении выпадают разного рода ошибки в буквах')
print('В нашем случае - поиск родственников по ИП')
print('')

# %%
non_mathes_1_6 = non_mathes_1_5.drop(potential_matches6['level_0'].array)
if 'index' in non_mathes_1_6.columns:
    non_mathes_1_6 = non_mathes_1_6.drop('index',axis=1)
non_fc_matches_1_6 = non_fc_matches_1_5.drop(potential_matches6['level_1'].array).reset_index(drop=True)

if len(potential_matches6)>0:
    non_matches_fin_4 = potential_matches6.apply(get_matches, df_lev_0=non_mathes_1_5, df_lev_1=non_fc_matches_1_5, axis=1).reset_index(drop=True)
    non_matches_fin_4.columns = non_mathes_1_5.columns.append(non_fc_matches_1_5.columns)
    if 'index' in non_matches_fin_4.columns:
        non_matches_fin_4 = non_matches_fin_4.drop('index', axis=1)
    process_data(non_matches_fin_4)
    non_matches_fin_4['Несоответствие'] = 'Совпадение по ИП > 80%'
else:
    non_matches_fin_2 = pd.DataFrame()

# %%
# Объединение неполных соответствий в 1 датафрейм
combined_non_matches = pd.concat([non_matches_fin_0, non_matches_fin_1, non_matches_fin_2, non_matches_fin_3, non_matches_fin_4], ignore_index=True)

# %% [markdown]
# ##### Вывод в файл пустых значений (критичных) и неполных соответствий
# 
# так же я могу добавить сюда 1 файл для записей с совпадающим должником и взыскателем, но разными ИД и ИП\
# но он работает, скажем так, долго, и к тому же на 35тыс оставшихся должников находит 100+тыс возможных соответствий\
# короче ситуация аналогична полным соответствиям - либо ручками, либо особо замудрёным алгоритмом, либо нужны доп параметры (последнее предпочтительно)\
# Если вам это не нужно - просто не запускайте блок ниже

# %%
print('Другие частичные соответствия искать я не вижу смысла, но если что -пишите')
print('')
print('И да - частичные соответствия - это соответствия после сверки между отчетом и реестром')
print('но данные совпадают в двух из четырех столбцов - в остальних, лишь частично')
print('неполные соответствия в свою очередь - строки, для которых соответствий не нашлось')
print('-------------------------------------------------------------------------------------------------------------------------------------------------------')
print('')

# %%
with pd.ExcelWriter('Пустые данные.xlsx') as writer:
    empty_data_292_debtor.to_excel(writer, sheet_name='Пустые должники в отчете', index=False, engine='openpyxl', startrow=0, startcol=0)
    empty_data_FCCP_debtor.to_excel(writer, sheet_name='Пустые должники в реестре', index=False, engine='openpyxl', startrow=0, startcol=0)
    empty_data_292.to_excel(writer, sheet_name='Пустые ИД и ИП в отчете', index=False, engine='openpyxl', startrow=0, startcol=0)
    empty_data_FCCP.to_excel(writer, sheet_name='Пустые ИД и ИП в реестре', index=False, engine='openpyxl', startrow=0, startcol=0)
    combined_ed_matches.to_excel(writer, sheet_name='Соотв. между пустыми данными', index=False, engine='openpyxl', startrow=0, startcol=0)
    combined_non_matches.to_excel(writer, sheet_name='Частичные соответствия', index=False, engine='openpyxl', startrow=0, startcol=0)

# %% [markdown]
# Финальная проверка- поиск нигде не задействованных строк

# %% [markdown]
# По отчету

# %%
print('Финальная проверка- поиск нигде не задействованных строк')
print('По отчету')
print('Всего строк после чистки от (критичных) пустых:',len(filtered_df292_2))
print('Всего строк с ненайденными соответствиями:',len(non_mathes_1_6))
print('Всего строк с найденными соответствиями:',len(matches_all_1))
print('Уникальных строк с найденными соответствиями:',matches_all_1['ID ИП'].nunique())
print('')
print('Всего строк с некорректными взыскателями (поиск по ИП):',len(non_matches_fin_1))
print('Уникальных строк с некорректными взыскателями (поиск по ИП):',non_matches_fin_1['ID ИП'].nunique())
print('')
print('Всего строк с некорректными взыскателями (поиск по ИД):',len(non_matches_fin_2))
print('Уникальных строк с некорректными взыскателями (поиск по ИД):',non_matches_fin_2['ID ИП'].nunique())
print('')
print('Всего найденных соответствий запросу (Совпадение по ИД > 80%):',len(non_matches_fin_3))
print('Уникальных соответствий запросу (Совпадение по ИД > 80%):',non_matches_fin_3['ID ИП'].nunique())
print('')
print('Всего найденных соответствий запросу (Совпадение по ИП > 80%):',len(non_matches_fin_4))
print('Уникальных соответствий запросу (Совпадение по ИП > 80%):',non_matches_fin_4['ID ИП'].nunique())
print('')
print('Всего строк с некорректными датами:',len(non_matches_fin_0))
print('Уникальных строк с некорректными датами:',non_matches_fin_0['ID ИП'].nunique())
print('')

print('Сравнение числа срок в начальном (отфильтрованном) отчете с суммой разбитых на категории')
print(len(filtered_df292_2), len(non_mathes_1_3)+ len(matches_all_1)+ len(non_matches_fin_0)+ len(non_matches_fin_1)+ len(non_matches_fin_2)+ len(non_matches_fin_3)+ len(non_matches_fin_4))
print('Разница в строках:', len(filtered_df292_2)-len(non_mathes_1_3)-len(matches_all_1)-len(non_matches_fin_0)-len(non_matches_fin_1)-len(non_matches_fin_2)- len(non_matches_fin_3)- len(non_matches_fin_4))
print('')
print('-------------------------------------------------------------------------------------------------------------------------------------------------------')

# %% [markdown]
# По реестру

# %%
print('По реестру')
print('Всего строк после чистки от (критичных) пустых:',len(dfFCCP_2))
print('Всего строк с ненайденными соответствиями:',len(non_fc_matches_1_6))
print('Всего строк с найденными соответствиями:',len(matches_all_1))
print('Уникальных строк с найденными соответствиями:',matches_all_1['ID ИП'].nunique())
print('')
print('Всего строк с некорректными взыскателями (поиск по ИП):',len(non_matches_fin_1))
print('Уникальных строк с некорректными взыскателями (поиск по ИП):',non_matches_fin_1['ID ИП'].nunique())
print('')
print('Всего строк с некорректными взыскателями (поиск по ИД):',len(non_matches_fin_2))
print('Уникальных строк с некорректными взыскателями (поиск по ИД):',non_matches_fin_2['ID ИП'].nunique())
print('')
print('Всего найденных соответствий запросу (Совпадение по ИД > 80%):',len(non_matches_fin_3))
print('Уникальных соответствий запросу (Совпадение по ИД > 80%):',non_matches_fin_3['ID ИП'].nunique())
print('')
print('Всего найденных соответствий запросу (Совпадение по ИП > 80%):',len(non_matches_fin_4))
print('Уникальных соответствий запросу (Совпадение по ИП > 80%):',non_matches_fin_4['ID ИП'].nunique())
print('')
print('Всего строк с некорректными датами:',len(non_matches_fin_0))
print('Уникальных строк с некорректными датами:',non_matches_fin_0['ID ИП'].nunique())

print('Сравнение числа срок в начальном (отфильтрованном) отчете с суммой разбитых на категории')
print(len(dfFCCP_2), len(non_fc_matches_1_3)+ len(matches_all_1)+ len(non_matches_fin_0)+ len(non_matches_fin_1)+ len(non_matches_fin_2)+ len(non_matches_fin_3)+ len(non_matches_fin_4))
print('Разница в строках:', len(dfFCCP_2)-len(non_fc_matches_1_3)-len(matches_all_1)-len(non_matches_fin_0)-len(non_matches_fin_1)-len(non_matches_fin_2)- len(non_matches_fin_3)- len(non_matches_fin_4))
print('')
print('-------------------------------------------------------------------------------------------------------------------------------------------------------')

# %% [markdown]
# Итог следующий - на данный момент я не могу избавиться от дубликатов - по двум блокам выше видно, сколько вы их можете найти\
# Для того, чтобы я смог обработать всё это должным образом нужно следующее:
# 1. Дополнительные параметры (столбцы, с одинаковыми значениями в обеих эксель, желательно сравнительно корректно заполненные)
# 2. Можно прописать то, как вы выделяете в итоге что из тех или иных дубликатов является корректным\
# другими словами - написанная последовательность действий, которую можно реализовать,\
#  без последующего задействования человеческого фактора для "посмотреть, что верно"
# 
# По большей части всё, больше я с текущими параметрами запроса ничего не выжму - либо редактируйте данные в базе,\
# чтобы соответствия всегда были явными, либо присылайте вышеуказанное для доработки.
# 
# А так удачи!)

# %%
print('Итог следующий - на данный момент я не могу избавиться от дубликатов - по двум блокам выше видно, сколько вы их можете найти')
print('Для того, чтобы я смог обработать всё это должным образом нужно следующее:')
print('1. Дополнительные параметры (столбцы, с одинаковыми значениями в обеих эксель, желательно сравнительно корректно заполненные)')
print('2. Можно прописать то, как вы выделяете в итоге что из тех или иных дубликатов является корректным')
print('другими словами - написанная последовательность действий, которую можно реализовать,')
print('без последующего задействования человеческого фактора для "посмотреть, что верно"')
print('')
print('По большей части всё, больше я с текущими параметрами запроса ничего не выжму - либо редактируйте данные в базе,')
print('чтобы соответствия всегда были явными, либо присылайте вышеуказанное для доработки.')
print('')
print('А так удачи!)')

# %% [markdown]
# #### Блок форматирования
# 
# non_mathes_1_4 - Неполные соответствия(отчет 292)\
# non_fc_matches_1_4 - Ненайденные соответствия(реестр ФССП)\
# matches_all_2 - Найденные соответствия
# 
# В случае, если вы хотите удалить ещё какие-нибудь столбцы - в соответствующих скобках добавьте их через запятую в формате 'Название столбца'

# %%
non_mathes_1_4 = non_mathes_1_3.loc[:, ~non_mathes_1_3.columns.isin(['index', 'level_0',  'ИП формат 292',  'ИД формат 292', 'Должник 292', 'Взыскатель'])]
non_fc_matches_1_4 = non_fc_matches_1_3.loc[:, ~non_fc_matches_1_3.columns.isin(['index', 'ИП формат ФССП', 'ИД формат ФССП', 'Должник ФССП', 'Взыскатель2', 'Взыскатель', 'Level_1'])]
matches_all_2 = matches_all_1.loc[:, ~matches_all_1.columns.isin(['index', 'level_0', 'ИП формат ФССП', 'ИП формат 292',  'ИД формат 292',
                                                                  'ИД формат ФССП', 'Должник 292', 'Должник ФССП', 'Взыскатель2', 'Взыскатель', 'Level_1'])]

# %% [markdown]
# #### Вывод в файл
# в первой строке - название файла (шаблон - 'результат.xlsx')
# в остальных - название листов
# 
# можно свободно переименовывать, главное не забывайте прописывать расширение файла и скобки ''

# %%
with pd.ExcelWriter('Результат поиска соответствий.xlsx') as writer:
    matches_all_2.to_excel(writer, sheet_name='Найденные соответствия', index=False, engine='openpyxl', startrow=0, startcol=0)
    non_fc_matches_1_6.to_excel(writer, sheet_name='Ненайденные соответствия(ФССП)', index=False, engine='openpyxl', startrow=0, startcol=0)
    non_mathes_1_6.to_excel(writer, sheet_name='Неполные соответствия(о292)', index=False, engine='openpyxl', startrow=0, startcol=0)

# %%
pd.set_option('display.max_rows', None)
#!{sys.executable} -m pip install nbformat -q
import nbformat
import os

with open('report_292_FCCP.ipynb', 'r', encoding='utf-8') as f:
    notebook = nbformat.read(f, as_version=4)

# Получаем все ячейки с выходами и объединяем их
output_text = ''
for cell in notebook.cells:
    if 'outputs' in cell and cell['cell_type'] == 'code':
        for output in cell['outputs']:
            if output.output_type == 'stream' and output.name == 'stdout':
                output_text += output.text

# Сохраняем выходные данные в файл
with open('Текстовые выводы внутри программы.txt', 'w', encoding='utf-8') as out_file:
    out_file.write(output_text)
if os.name == 'nt': # Проверяем, если операционная система Windows
    os.system('pause')


