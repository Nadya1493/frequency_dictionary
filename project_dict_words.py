import string
from collections import defaultdict
import nltk
import pymorphy2
import xlwt
from nltk.corpus import stopwords

# открытие и чтение файла
way_txt = input('Введите путь к файлу: ')
file = open(way_txt, "r", encoding="UTF-8")
text = file.read()
text = text.lower()

# создание словаря "слово-количество"
dict_words_count = defaultdict(int)

# разбить текст на предложения
sentences = nltk.sent_tokenize(text)

# работа с отдельными предложениями, удаление знаков препинания,, кавычек, стоп-слов
stop_words = list(stopwords.words("russian"))
stop_words.extend(['что', 'это', 'так', 'вот', 'быть', 'как', 'в', '—', '–', 'к', 'на', '...', 'б', 'бы'])

for sentence in sentences:
    # разбить предложения на слова
    words = nltk.word_tokenize(sentence)
    # удаление знаков препинания
    words = [i for i in words if (i not in string.punctuation)]
    # удаление кавычек
    words = [i.replace("«", "").replace("»", "") for i in words]
    # удаление стоп слов
    text_without_stop_words = [word for word in words if (word not in stop_words)]
    # приведение слов к нормальной форме, добавление "слово-количество" в словарь
    for text_without_stop_word in text_without_stop_words:
        word_morph = pymorphy2.MorphAnalyzer().parse(text_without_stop_word)[0].normal_form
        dict_words_count[word_morph] += 1

# сортировка словаря по значениям в порядке убывания
list_dict_items = list(dict_words_count.items())
list_dict_items.sort(key=lambda i: i[1], reverse=True)

# запись и сохранение в файл Exel
wb = xlwt.Workbook()
ws = wb.add_sheet('Words')
count = 0
for k, v in list_dict_items:
    ws.write(count, 0, k)
    ws.write(count, 1, v)
    count += 1

wb.save(r'.\xl_rez.xls')

file.close()
