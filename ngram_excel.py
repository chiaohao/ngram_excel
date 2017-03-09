# -*- coding: utf-8 -*-
import xlrd
import xlwt
import codecs
import operator
import sys

cutlist = "1234567890<>/:：;；,、＂’，.。！？「」（）｢\"\'\\\n\r《》“”!@#$%^&*()".decode("utf-8")

def cutSentence(workbook_path, keywords, contentColumnIndex): 
	in_workbook = xlrd.open_workbook(workbook_path)
	text_column = in_workbook.sheets()[0].col_values(contentColumnIndex)
	row_counts = in_workbook.sheets()[0].nrows
	#print repr([a.encode(sys.stdout.encoding) for a in text_column]).decode("string-escape")
	sentence = ""
	textList = []
	count = 0
	for line in text_column:
		line = line.strip()
		
		for keyword in keywords:
			line = "".join(line.split(keyword))
		
		for word in line:
			if word not in cutlist:
				sentence += word
			else:
				textList.append(sentence)
				sentence = ""
		count += 1
		if count == row_counts:
			break
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

def longTermPriority(path, maxTermLength, minFreq, contentColumnIndex):
	longTerms=[]
	longTermsFreq=[]
	
	for i in range(maxTermLength,1,-1):
		text_list = cutSentence(path,longTerms, contentColumnIndex)
		words_freq = ngram(text_list,i, minFreq)
	
		for word_freq in words_freq:
			longTerms.append(word_freq[0])
			longTermsFreq.append(word_freq) 
	
	return longTermsFreq

#python ngram_excel.py in_workbook content_column out_workbook longest_gram_num min_freqency
#                      sys.argv[1] sys.argv[2]    sys.argv[3]  sys.argv[4]      sys.argv[5]

col = ord(sys.argv[2].lower()) - ord('a')

longTermFreq = longTermPriority(sys.argv[1], int(sys.argv[4]), int(sys.argv[5]), col)

out_workbook = xlwt.Workbook()
out_table = out_workbook.add_sheet("output")

c = 0
for i in longTermFreq:
	out_table.write(c, 0, i[0])
	out_table.write(c, 1, i[1])
	c = c + 1

out_workbook.save(sys.argv[3])

