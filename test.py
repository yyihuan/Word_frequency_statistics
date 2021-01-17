from 词频统计查询工具 import *
import os

cwd = os.getcwd()
bookname = "test.xlsx"
book_io = cwd + "/" + bookname  # 对应Book的excel文件的绝对路径
wordlist_name = "test_leadin.txt"
wordlist_io = r"{}/{}".format(cwd, wordlist_name)

pte = Book(book_io)
print(pte.Sheet1.show())
pte.Sheet1.lead_in(wordlist_io)
print(pte.Sheet1.show())