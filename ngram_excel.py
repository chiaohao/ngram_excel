# -*- coding: utf-8 -*-
import codecs
import operator
import sys
import xlrd
import xlsxwriter

cutlist = " <>/:：;；,、＂’，.。！？「」（）｢\"\'\\\n\r《》“”!@#$%^&*()".decode("utf-8")

def cutSentence(text_column): 
	#print repr([a.encode(sys.stdout.encoding) for a in text_column]).decode("string-escape")
	sentence = ""
	textList = []

	for line in text_column:
		line = line.strip()
		
		for word in line:
			if word not in cutlist:
				sentence += word
			else:
				textList.append(sentence)
				sentence = ""

	return textList

def ngram(textLists,n,minFreq): 
	words_freq={}
	result= []
	for textList in textLists:
		for w in range(len(textList)-(n-1)):
			word = textList[w:w+n]
			if word not in words_freq:
				words_freq[word] = 1
			else:
				words_freq[word] += 1

	for word in words_freq:
		if words_freq[word] >= minFreq:
			result.append([word, words_freq[word]])

	return result

def longTermPriority(text_column, maxTermLength, minFreq):
	longTermsFreq=[]
	
	for i in range(maxTermLength,1,-1):
		text_list = cutSentence(text_column)
		words_freq = ngram(text_list,i, minFreq)
	
		for word_freq in words_freq:
			longTermsFreq.append(word_freq) 
	
	return longTermsFreq

def CountDocumentFrequency(text_column, gram_with_tf):
	gram_tf_df = []

	for gram in gram_with_tf:
		dfCount = 0

		for doc in text_column:
			if gram[0] in doc:
				dfCount += 1

		gram_tf_df.append([gram[0], gram[1], dfCount])

	return gram_tf_df

def gramDF(gram, text_column):


#python ngram_excel.py in_workbook content_column out_workbook longest_gram_num min_freqency
#                      sys.argv[1] sys.argv[2]    sys.argv[3]  sys.argv[4]      sys.argv[5]


col = ord(sys.argv[2].lower()) - ord('a')

in_workbook = xlrd.open_workbook(sys.argv[1])
text_column = in_workbook.sheets()[0].col_values(col)
row_counts = in_workbook.sheets()[0].nrows

longTermFreq = longTermPriority(text_column, int(sys.argv[4]), int(sys.argv[5]))
print "tf done"
gram_tf_df = CountDocumentFrequency(text_column, longTermFreq)
print "df done"

out_workbook = xlsxwriter.Workbook(sys.argv[3])
out_table = out_workbook.add_worksheet()

out_table.write(0, 1, "docNum")
out_table.write(0, 0, row_counts)

out_table.write(2, 0, "gram")
out_table.write(2, 1, "tf")
out_table.write(2, 2, "df")

c = 3
for i in gram_tf_df:
	out_table.write(c, 0, i[0])
	out_table.write(c, 1, i[1])
	out_table.write(c, 2, i[2])
	c = c + 1

out_workbook.close()
