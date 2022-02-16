from bs4 import BeautifulSoup
import pandas as pd

source_path = 'page.html'
result_path = 'marks.xlsx'


with open(source_path, 'r') as sf:
    code = sf.read()

soup = BeautifulSoup(code, 'html.parser')

# Get target tables from source file
tables = soup.find_all('div', {'class': 'table-responsive scrolltable-ready table-light'})

if (not tables):
    print('Check your file content (no tables)')
    exit()


# Prepare using data
result = pd.DataFrame()
target_headers = ['Семестр', 'Дисциплина', 'Кол-во часов', 'Объем ЗЕ', 'Вид контроля', 'Попытка', 'Отметка']

sem = 0

for table in tables:
    sem_flag = 0
    
    for row in table.find_all('tr'):
        # If row contain headers
        if (row.find('th')):
            # And contain semestr name
            if (row.find('th').get('class') == ['apex_report_break']):
                sem +=1
            else:
                # Add semestr column if not exist
                av_headers = [col.text for col in row.find_all('th')]
                if ('Семестр' not in av_headers):
                    av_headers = ['Семестр'] + av_headers
                    sem_flag = 1
            continue

        data = [sem] if sem_flag else []
        cols = row.find_all('td')

        # Compare target and available columns
        for i in range(sem_flag, len(av_headers)-1):
            if (av_headers[i] in target_headers):
                data.append(cols[i-sem_flag].text)
                
        # Use just matched columns
        headers = [head for head in av_headers if head in target_headers]

        result = result.append(pd.DataFrame([data], columns=headers), ignore_index=True)

result.to_excel('result.xlsx')