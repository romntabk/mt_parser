import os 
import pandas as pd 
from bs4 import BeautifulSoup


''' EXCEL INFO '''
EXCEL_FILE_NAME = 'RU-EXCEL.xlsx' # название Excel файла 
PATH_TO_EXCEL = '' # путь к Excel файлу
SHEET_NAME = '10.12.2021' # название страницы в Excel файле
MARK = '♫' # метка по которой идёт поиск


''' XML INFO '''
PATH_TO_XML_FILES = 'xml_folder/' 
RESULT_FILE_NAME = EXCEL_FILE_NAME.split('.')[0].replace('EXCEL','XML')+'.xml' # название итогового файла


''' DATA '''
INDEXES = {} # номера столбцов 'play','артист','композиция'
ARTISTS = [] # список нужных артистов



'''
	Ищет номера столбцов 'play','артист','композиция' 
	и заполняет словарь INDEXES ими. Для случая изменения 
	структуры таблицы excel.
'''
def set_indexes(df):
	for index, row in df.iterrows():
		row_lower = [i.lower() if type(i) is type('') else '' for i in row]  # приводит к нижнему регистру
		if 'play' not in row_lower:
			continue
		assert 'артист' in row_lower, 'Нет столбца артист' # выдать ошибку если не нашли 'артист', но нашли PLAY
		assert 'композиция' in row_lower, 'Нет столбца композиция' # выдать ошибку если не нашли 'композиция', но нашли PLAY
		global INDEXES
		INDEXES = {i:row_lower.index(i) for i in ['play','артист','композиция']}
		return
	assert False, 'Нет столбца PLAY' # выдать ошибку если не нашли play 



'''
	Заполняет список ARTISTS словарями
	вида {'имя': имя, 'композиция': композиция}
'''
def find_excel_artists(df):
	for index, row  in df.iterrows():
		if row[INDEXES['play']] == MARK:
			ARTISTS.append({'имя' : row[INDEXES['артист']],'композиция' : row[INDEXES['композиция']]})


'''
	Ищет все нужные clipitems в xml_file
	и сразу записывает в файл
'''
def find_clipitems(xml_file):
	file = open(PATH_TO_XML_FILES+xml_file,encoding='utf-8')
	xml= file.read()
	soup = BeautifulSoup(xml,'lxml')
	file_writer = open(RESULT_FILE_NAME, encoding='utf-8', mode='w')
	count=0
	for tag in soup.findAll('clipitem'):
		result = check_clipitem(tag)
		if result:
			file_writer.write(tag.prettify())
			count+=1
	print(f'В файле {xml_file} найдено {count} clipitem\'ов')



'''
	tag - clipitem в котором идёт поиск артиста
	возвращает имя артиста, если найден,
	False, если не найден
'''
def check_clipitem(tag):
	for artist in ARTISTS:
		if artist['имя'] in str(tag.find('name')):
			return artist['имя']
	return False


def parse():
	t = pd.ExcelFile(PATH_TO_EXCEL+EXCEL_FILE_NAME)
	df = t.parse(SHEET_NAME)
	set_indexes(df)
	find_excel_artists(df)
	for file_name in os.listdir(PATH_TO_XML_FILES):
		find_clipitems(file_name)


if __name__ == '__main__':
	parse()
