import string
from collections import Counter  # for counting elems and n-grams

import docx  # docx
import fitz  # pdf
import nltk  # language processing
import pandas as pd  # csv and xls
import pymorphy2  # lemmatization
from matplotlib import pyplot as plt
from nltk import ngrams  # ngrams
from nltk import word_tokenize  # tokenization
from nltk.corpus import stopwords  # deleting spam
from nltk.probability import FreqDist  # distribution


def open_source_file(file_name):  # открывает исходный файл
    if file_name.endswith('txt'):
        source_file_data = open(f'{file_name}', "r", encoding='UTF-8')  # указываем открываемый файл и характеристики
        raw_source_text = source_file_data.read()
        return raw_source_text
    elif file_name.endswith('pdf'):
        pdf_document = file_name
        doc = fitz.open(pdf_document)
        text = []  # формируем список текста в страницах
        pages = range(doc.page_count)
        for page in pages:
            page = doc.load_page(page)
            text.append(page.get_text("text"))
        raw_source_text = "\n".join(text)
        return raw_source_text
    elif file_name.endswith('docx'):
        source_file_data = docx.Document(f'{file_name}')
        text = []  # формируем список текста в абзацах
        for paragraph in source_file_data.paragraphs:
            text.append(paragraph.text)
        raw_source_text = "\n".join(text)
        return raw_source_text


def raw_text_processing(raw_source_text):  # избавляет текст от пунктуации
    raw_text_length = len(raw_source_text)
    raw_source_text = raw_source_text.lower()
    spec_char = string.punctuation + '\n\xa0«»\t—–-…“„'  # добавляем нужные символы
    raw_source_text = "".join([char for char in raw_source_text if char not in spec_char])  # нет пунктуации
    cleared_text = "".join([char for char in raw_source_text if char not in string.digits])  # нет чисел
    return cleared_text, raw_text_length


def tokenization(cleared_text):  # выделяет слова из очищенной "массы"
    text_tokens = word_tokenize(cleared_text)
    tokenized_text = nltk.Text(text_tokens)
    return tokenized_text, text_tokens


def spam_words_delete(text_tokens):  # удаляет стоп-слова, слова паразиты и все бесполезное
    russian_stopwords = stopwords.words("russian")
    russian_stopwords.extend(['это', 'нею', 'б', 'ах'])  # добавляем в список стопслов свои
    no_spam_tokens = []
    for token in text_tokens:
        if token not in russian_stopwords:
            no_spam_tokens.append(token)
    return no_spam_tokens


def lemmatization(no_spam_tokens):  # приводит слова к нормальной, "словарной" форме
    lemmatized_tokens = []
    single_lemmatized_tokens = []
    morph = pymorphy2.MorphAnalyzer()
    for word in no_spam_tokens:  # запускаем проверку слов
        parsed_word = morph.parse(word)[0]
        lemmatized_tokens.append(parsed_word.normal_form)
    for word in lemmatized_tokens:
        if word not in single_lemmatized_tokens:
            single_lemmatized_tokens.append(word)
    return lemmatized_tokens, single_lemmatized_tokens


def storing_to_dataframe(lemmatized_tokens, single_lemmatized_tokens):  # записываем результаты в таблицу
    lemmatized_data = dict(col1=lemmatized_tokens)
    single_lemmatized_data = dict(col1=single_lemmatized_tokens)
    single_lemmatized_dataframe = pd.DataFrame(single_lemmatized_data)
    single_lemmatized_dataframe.to_excel(r'F:\project\Список слов без повторений.xlsx', index=False)
    lemmatized_dataframe = pd.DataFrame(lemmatized_data)
    lemmatized_dataframe.to_excel(r'F:\project\Список слов.xlsx', index=False)
    lemmatized_dataframe.to_csv(r'F:\project\Список слов.csv', encoding='UTF-8', index=False)


def get_most_common_words(text):
    freq_of_dist = FreqDist(text)
    most_common_word = (freq_of_dist.most_common(1))[0]
    return most_common_word, freq_of_dist


def frequency_plotter(freq_of_dist, text):
    freq_of_dist = dict(freq_of_dist)  # создаем таблицу со всеми словами и частотой употребления
    sorted_freq_of_dist = {}
    sorted_keys = sorted(freq_of_dist, key=freq_of_dist.get, reverse=True)
    for w in sorted_keys:
        sorted_freq_of_dist[w] = freq_of_dist[w]
    f = list(sorted_freq_of_dist.items())[:30]
    first30words = []
    for pair in f:
        first30words.append(pair[0])
        first30words.append(pair[1])
    x = [v for k, v in enumerate(first30words) if not k % 2]
    y = [v for k, v in enumerate(first30words) if k % 2]
    plt.xlabel("Слова")
    plt.ylabel("Количество употреблений")
    plt.grid()
    plt.plot(x, y)
    plt.xticks(rotation=90)
    plt.savefig('1.png')


def ngrams_cal(text_input):
    i = int(input('1 - провести подсчет для слов\n'
                  '2 - провести подсчет для символов\n'))
    if i == 1:
        n = int(input('Для скольки n провести рассчет? '))
        raw_grams = ngrams(text_input, n)
        counted_ngrams = Counter(raw_grams)
        df = pd.DataFrame(counted_ngrams.keys())
        df["Количество употреблений"] = counted_ngrams.values()
        df.to_excel(r'F:\project\Результат подсчета n-грамм.xlsx')
    elif i == 2:
        united_text = ''.join(text_input)
        n = int(input('Для скольки n провести рассчет? '))
        raw_grams = ngrams(united_text, n)
        counted_ngrams = Counter(raw_grams)
        df = pd.DataFrame(counted_ngrams.keys())
        df["Количество употреблений"] = counted_ngrams.values()
        df.to_excel(r'F:\project\Результат подсчета n-грамм.xlsx')


def main():
    def data_to_docx():
        data_out = docx.Document()
        data_out.add_paragraph("Характеристики текста:")
        data_out.add_paragraph(f'Символов до обработки: {len(raw_source_text)}')
        data_out.add_paragraph(f'Символов после удаления пунктуации: {len(cleared_text)}')
        data_out.add_paragraph(f'Знаков пунктуации: {len(raw_source_text) - len(cleared_text)}')
        data_out.add_paragraph(f'Текст содержит слов: {len(text_tokens)}')
        data_out.add_paragraph(f'Самое частое слово до уборки стоп-слов:'
                               f' "{most_common_word[0]}" в количестве {most_common_word[1]}')
        data_out.add_paragraph(f'Самое частое слово после уборки стоп-слов:'
                               f'"{cleared_most_common_word[0]}" в количестве {cleared_most_common_word[1]}')
        data_out.add_paragraph(f'Слов после удаления стоп-слов: {len(no_spam_tokens)}')
        data_out.add_paragraph(f'Удалено стоп-слов: {len(text_tokens) - len(no_spam_tokens)}')
        data_out.add_paragraph(f'При лемматизации удалено слов: {len(no_spam_tokens) - len(single_lemmatized_tokens)}')
        data_out.add_paragraph(f'Уникальных слов: {len(single_lemmatized_tokens)}')
        data_out.add_paragraph("Графики:")
        frequency_plotter(freq_of_dist, tokenized_text)
        data_out.add_paragraph("График 1. Слова до уборки стоп-слов")
        data_out.add_picture('1.png')
        frequency_plotter(cleared_freq_of_dist, no_spam_tokens)
        data_out.add_paragraph("График 2. Слова после уборки стоп-слов")
        data_out.add_picture('1.png')
        data_out.save(r"F:\project\Результаты анализа.docx")

    def get_input_ngrams():
        w = int(input('1 - посчитать n-граммы для чистого текста без стоп-слов\n'
                      '2 - посчитать n-граммы для лемматизованного текста\n'
                      '3 - посчитать n-граммы для текста без пунктуации\n'
                      '4 - посчитать n-граммы для исходного текста\n'))
        if w == 1:
            ngrams_cal(no_spam_tokens)
        elif w == 2:
            ngrams_cal(lemmatized_tokens)
        elif w == 3:
            ngrams_cal(tokenized_text)
        elif w == 4:
            ngrams_cal(raw_source_text)

    file_name = r"F:\project\1.txt"
    raw_source_text = open_source_file(file_name)
    cleared_text = raw_text_processing(raw_source_text)[0]
    tokenized_text, text_tokens = tokenization(cleared_text)
    no_spam_tokens = spam_words_delete(text_tokens)
    lemmatized_tokens, single_lemmatized_tokens = lemmatization(no_spam_tokens)
    storing_to_dataframe(lemmatized_tokens, single_lemmatized_tokens)
    most_common_word, freq_of_dist = get_most_common_words(tokenized_text)
    cleared_most_common_word, cleared_freq_of_dist = get_most_common_words(no_spam_tokens)

    data_to_docx()
    print('Результаты анализа занесены в файлы')
    get_input_ngrams()


if __name__ == "__main__":
    main()
