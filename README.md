# Ngram Analysis Excel Input/Output
Read a column of excel workbook and return a sheet with n-grams and counts

# Install
> pip install xlrd
> pip install xlwt
> git clone https://github.com/chiaohao/ngram_excel.git
> cd ngram_excel

#Usage
``` python ngram_excel.py in_workbook content_column out_workbook longest_gram_num min_freqency
in_workbook: your input workbook name
content_column: column that you want to analysis, for example, 'a', 'd'...
out_workbook: your output workbook name (remember to add .xls)
longest_gram_num: for example, '6' means 2 to 6 grams
min_freqency: minimun count of a gram to output