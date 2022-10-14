import pandas as pd
import nltk
from openpyxl import load_workbook
# Сте́мминг — это процесс нахождения основы слова для заданного исходного слова. Основа слова не обязательно совпадает с морфологическим корнем слова.
from nltk import word_tokenize
from nltk.stem.snowball import SnowballStemmer  # библиотека для самого стеммера
# Стоп-слова - дополнительные слова, которые не несут смысловой нагрузки. К ним относятся местоимения, частицы и некоторые общеупотребительные глаголы.
from nltk.corpus import stopwords
from nltk.stem import *

stemmer = SnowballStemmer("russian")
nltk.download('stopwords')
nltk.download('punkt')
# тут теперь лежат все стоп-слова
russian_stopwords = stopwords.words("russian")
russian_stopwords.extend(['…', '«', '»', '...', 'здравствуйте', 'который', 'это', 'пожалуйста',\
                         'спасибо', 'этот', 'наш', 'никто', 'свой', 'т.д.', 'т', 'д'])  # добавил парочку от себя


wb = load_workbook(filename='Courses.xlsx')
ws = wb.active

d = {}
for row in list(ws.rows)[1:]:
    d[row[0].value] = [str(c.value) for c in row[1:]]

# тут начинается стемминг
stemmed_texts_list = []
for i, text in d.items():  # проходимся по всему словарю курсов-описаний
    # разбиваем текст на токены. Токенизация – процесс разбиения текста на текстовые единицы, например, слова или предложения.
    tokens = word_tokenize(text[0])
    stemmed_tokens = [stemmer.stem(\
        token) for token in tokens if token not in russian_stopwords]  # убираем стоп-слова
    text = " ".join(stemmed_tokens)  # возвращаем в текст
    d[i] = text  # передаем стеммингованный текст обратно в словарь
    stemmed_texts_list.append(text)  # добавляем весь текст в общий список

# тут будем искать самые часто встречающиеся слова
words = {}
for i in stemmed_texts_list:
    for j in i.split():
        # создаем словарь частоты слов, ключ - слово, значение - сколько раз встретилось
        words[j] = words.get(j, 0) + 1

# тут тупо сортим словарь
sorted_values = sorted(words.values())
sorted_dict = {}
for i in sorted_values:
    for k in words.keys():
        if words[k] == i:
            sorted_dict[k] = words[k]
            break

# тут выбираем 20 самых частых слов и заносим их в стоп-слова
russian_stopwords.append(list(sorted_dict.keys())[-20:])




def matchingElements(dictionary, searchString):
    for key, val in dictionary.items():
        if searchString in val:
            return key
    return ''


# создаем новую таблицу для окончательного результата
new = pd.DataFrame(columns=['prof', 'comp', 'indc', 'course', 'description'])
tbl_courses = pd.read_excel('Courses.xlsx')
tbl = pd.read_excel('Новая таблица (1).xlsx', sheet_name='1')
for i in tbl.itertuples():  # проходимся по таблице индикаторов
    ind, prof, comp, indc = i
    # тут снова проводим стемминг для самого индикатора
    tokens = word_tokenize(indc)
    stemmed_tokens = [stemmer.stem(\
        token) for token in tokens if token not in russian_stopwords]
    text = " ".join(stemmed_tokens)
    courses = set()
    for i in text.split():  # для каждого слова в индикаторе ищем подходящий курс (имеется в виду, что в отобранных словах точно есть ключевое и убраны все не ключевые)
        courses.add(matchingElements(d, i))
    courses.discard('')  # тут я просто убрал пустые списки
    courses = list(courses)  # перегоняю в список
    for i in courses:
        # для каждого курса нахожу описание
        description = tbl_courses['Description'][tbl_courses[' Name'] == i].iloc[0]
        new = new.append({'prof': prof, 'comp': comp, 'indc': indc, 'course': i,\
                         'description': description}, ignore_index=True)  # добавляю стороку в исходную таблицу
new.to_excel('new.xlsx', index=False)  # сохраняю таблицу в файл
