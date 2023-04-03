import string
from collections import Counter  # for counting elms and n-grams
import docx  # docx
import fitz  # pdf
import nltk  # language processing
import pandas as pd  # csv and xls
import pymorphy2  # lemmatization
from tkinter import filedialog  # interface
from tkinter import *
from tkinter import messagebox
from tkinter.ttk import Checkbutton
from tkinter.ttk import Combobox
from matplotlib import pyplot as plt  # graphics
from nltk import ngrams  # ngrams
from nltk import word_tokenize  # tokenization
from nltk.corpus import stopwords  # deleting spam
from nltk.probability import FreqDist  # distribution


def open_source_file(file_name):  # открывает исходный файл
	if file_name.endswith('txt'):
		source_file_data = open(f'{file_name}', "r",
								encoding='UTF-8')  # указываем открываемый файл и характеристики
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


def storing_to_dataframe(lemmatized_tokens, single_lemmatized_tokens,
						 saved_file_name):  # записываем результаты в таблицу
	lemmatized_data = dict(col1=lemmatized_tokens)
	single_lemmatized_data = dict(col1=single_lemmatized_tokens)
	single_lemmatized_dataframe = pd.DataFrame(single_lemmatized_data)
	single_lemmatized_dataframe.to_excel(f'{saved_file_name}\\Список слов без повторений.xlsx', index=False)
	lemmatized_dataframe = pd.DataFrame(lemmatized_data)
	lemmatized_dataframe.to_excel(f'{saved_file_name}\\Список слов.xlsx', index=False)


def get_most_common_words(text):
	freq_of_dist = FreqDist(text)
	most_common_word = (freq_of_dist.most_common(1))[0]
	return most_common_word, freq_of_dist


def frequency_plotter(freq_of_dist):
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
	plt.autoscale()
	plt.savefig('1.png')
	plt.clf()


def get_file_and_directory():
	messagebox.showinfo('Уведомление', 'Выберите файл для анализа')
	file_name = filedialog.askopenfilename()
	messagebox.showinfo('Уведомление', 'Выберите или создайте папку для сохранения результатов')
	saved_file_name = filedialog.askdirectory()
	return file_name, saved_file_name


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
		data_out.add_paragraph(
			f'При лемматизации удалено слов: {len(no_spam_tokens) - len(single_lemmatized_tokens)}')
		data_out.add_paragraph(f'Уникальных слов: {len(single_lemmatized_tokens)}')
		data_out.add_paragraph("Графики:")
		data_out.add_paragraph("График 1. Слова до уборки стоп-слов")
		frequency_plotter(freq_of_dist)
		data_out.add_picture('1.png')
		frequency_plotter(cleared_freq_of_dist)
		data_out.add_paragraph("График 2. Слова после уборки стоп-слов")
		data_out.add_picture('1.png')
		data_out.save(f'{saved_file_name}\\Результаты анализа.docx')

	file_name, saved_file_name = get_file_and_directory()
	raw_source_text = open_source_file(file_name)
	cleared_text = raw_text_processing(raw_source_text)[0]
	tokenized_text, text_tokens = tokenization(cleared_text)
	no_spam_tokens = spam_words_delete(text_tokens)
	lemmatized_tokens, single_lemmatized_tokens = lemmatization(no_spam_tokens)
	storing_to_dataframe(lemmatized_tokens, single_lemmatized_tokens, saved_file_name)
	most_common_word, freq_of_dist = get_most_common_words(tokenized_text)
	cleared_most_common_word, cleared_freq_of_dist = get_most_common_words(no_spam_tokens)

	def ngrams_cal(saved_file_name, text_input, ngram_count, word_or_symbol):
		if word_or_symbol == 1:
			united_text = ''.join(text_input)
			raw_grams = ngrams(united_text, ngram_count)
			counted_ngrams = Counter(raw_grams)
			df = pd.DataFrame(counted_ngrams.keys())
			df["Количество употреблений"] = counted_ngrams.values()
			df.to_excel(f'{saved_file_name}\\Результат подсчета n-грамм.xlsx')
		elif word_or_symbol == 2:
			raw_grams = ngrams(text_input, ngram_count)
			counted_ngrams = Counter(raw_grams)
			df = pd.DataFrame(counted_ngrams.keys())
			df["Количество употреблений"] = counted_ngrams.values()
			df.to_excel(f'{saved_file_name}\\Результат подсчета n-грамм.xlsx')

	return data_to_docx, ngrams_cal, saved_file_name, raw_source_text, tokenized_text, no_spam_tokens, lemmatized_tokens


def interface():
	def start():
		global window
		window.destroy()

		def ngramm_window_show():
			def choice():
				which_ngram = get_selected_type.get()
				ngram_count = int(how_many_ngram.get())
				word_or_symbol = get_word_or_symbol.get()

				def get_ngram_param(which_ngram, ngram_count, word_or_symbol):
					if which_ngram == 1:
						program[1](program[2], program[3], ngram_count, word_or_symbol)
					elif which_ngram == 2:
						program[1](program[2], program[4], ngram_count, word_or_symbol)
					elif which_ngram == 3:
						program[1](program[2], program[5], ngram_count, word_or_symbol)
					elif which_ngram == 4:
						program[1](program[2], program[6], ngram_count, word_or_symbol)
				get_ngram_param(which_ngram, ngram_count, word_or_symbol)
				ngram_window.destroy()

			ngram_window = Tk()
			ngram_window.title("Анализ n-грамм")
			ngram_window.geometry('500x250')
			get_selected_type = IntVar()
			get_word_or_symbol = IntVar()
			rad1 = Radiobutton(ngram_window, text='Оригинальный текст', variable=get_selected_type, value=1)
			rad2 = Radiobutton(ngram_window, text='Текст без пунктуации и цифр', variable=get_selected_type, value=2)
			rad3 = Radiobutton(ngram_window, text='Чистый текст без стоп-слов', variable=get_selected_type, value=3)
			rad4 = Radiobutton(ngram_window, text='Лемматизованный текст', variable=get_selected_type, value=4)
			rad5 = Radiobutton(ngram_window, text='Анализ символов в тексте', variable=get_word_or_symbol, value=1)
			rad6 = Radiobutton(ngram_window, text='Анализ слов в тексте', variable=get_word_or_symbol, value=2)
			btn = Button(ngram_window, text="Запустить анализ n-грамм", command=choice)
			how_many_ngram = Combobox(ngram_window)
			how_many_ngram['values'] = (2, 3, 4, 5, 6, 7, 8, 9, 10)
			how_many_ngram.current(0)
			rad1.grid(column=0, row=0)
			rad2.grid(column=0, row=1)
			rad3.grid(column=0, row=2)
			rad4.grid(column=0, row=3)
			rad5.grid(column=0, row=4)
			rad6.grid(column=1, row=4)
			how_many_ngram.grid(column=0, row=5)
			btn.grid(column=0, row=6)
			ngram_window.mainloop()

		check1 = chk_state1.get()
		check2 = chk_state2.get()
		if check1 == 1 or check2 == 1:
			program = main()
			if check1 == 1:
				program[0]()
				messagebox.showinfo('Уведомление', 'Этап базового анализа выполнен')
				messagebox.showinfo('Уведомление', 'Работа завершена \nМожно выходить')
			if check2 == 1:
				ngramm_window_show()
				messagebox.showinfo('Уведомление', 'Этап подсчета n-грамм выполнен')
				messagebox.showinfo('Уведомление', 'Работа завершена \nМожно выходить')
		elif not check1 != 1 and not check2 != 1:
			messagebox.showinfo('Уведомление', 'Вы не выбрали ни одного параметра')

	def show_guide():
		guide_text_source = open('D:\\PITON\\guide_text.txt', "r", encoding='UTF-8')
		guide_text = guide_text_source.read()
		guide_window = Tk()
		guide_window.title('Инструкция по применению')
		guide_window_label = Label(guide_window, text=guide_text, justify='left')
		guide_window_label.grid(column=0, row=0)

	global window
	window = Tk()
	window.title("Анализатор ЕЯ")
	window.geometry('500x250')
	label = Label(window, text="Анализатор естественного языка", font=("Arial Bold", 14), justify='right')
	label.grid(padx=0, pady=0, column=0, row=0)
	guide = Button(window, text="Инструкция по применению", command=show_guide)
	guide.grid(padx=0, pady=20, column=0, row=1)
	start_button = Button(window, text="ПУСК", command=start)
	start_button.grid(padx=0, pady=25, column=0, row=4)
	chk_state1 = IntVar()
	chk_state1.set(1)
	chk1 = Checkbutton(window, text='Базовый анализ', var=chk_state1)
	chk1.grid(column=0, row=2)
	chk_state2 = IntVar()
	chk_state2.set(1)
	chk2 = Checkbutton(window, text='Анализ n-грамм', var=chk_state2)
	chk2.grid(column=0, row=3)
	window.mainloop()


interface()
