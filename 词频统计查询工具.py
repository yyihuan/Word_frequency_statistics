# coding: utf-8

"""
本工具目的旨在统计已经查询过单词的查询次数

例如，我要准备PTE考试，那么我新建一个book叫PTE，它对应一个excel文件。那么不同题型的词汇，就是不同的sheet，即sheet1是FR之类。
我可以查看每个题型的词汇词频，也可以查看某几个sheet或整本书的单词词频。
    
结构：
     Book----sheet1
           |-sheet2
           |-······
           |-sheetN
    即每张词汇表为一个sheet实例，不同词汇表和归档为一个book实例。和Excel结构类似


暂定输出为：
    1. 程序内，查询或者增加次数后，显示单词以及频次
    2. 输出一个excel表格

后期计划为
    1. 可统计语料库（即文章集合）内词频（各词性形态是个问题。。。）
    2. pyqt改造GUI
    3. 导入词汇列表，并统计，然后更新原有列表 √
    4. 增加分组功能，分别统计  √
    5. 数据可视化
        1. 云图
    6. 增加保存格式，txt、csv等
"""

import pandas as pd
import os

cwd = os.getcwd()
bookname = "test.xlsx"
book_io = cwd + "/" + bookname  # 对应Book的excel文件的绝对路径
wordlist_name = "test_leadin.txt"
wordlist_io = r"{} / {}".format(cwd, wordlist_name)路径

class Book:
    
    def __init__(self, io):
        """
        将指定excel文件转换为数个单词表，每个单词表都是一个Sheet实例
        """
        self.sheet_list = [i for i in pd.read_excel(io, sheet_name=None).keys()]
        for i in self.sheet_list:
            exec("self.{} = Sheet(io, '{}')".format(i, i))
            print("{} sheet have been created".format(i))
    
    def del_sheet(self, sheet_list):
        """删除指定单词表
        
        Args:
        
            sheet_list: str/[str1, str2, ...]  单词表名
        """
        pass
    
    def create_sheet(self, sheet_list):
        """
        创建并保存新的单词表
        
        Args:
        
            sheet_list: str/[str1, str2, ...]  单词表名
        """
        pass


class Sheet:
    
    def __init__(self, io, sheet):
        """
        通过Book实例参数，读取已有表单文件内的表格，到df数据(单词表)。该DF数据即实际操作对象。
        
        单词表数据包括：单词-对应出现次数；总单词数；总共出现单词次数
        """
        self.df = pd.read_excel(io, sheet_name=0, index_col= "word")
        self.io = io
        self.sheet = sheet
        
    def set(self, word):
        """
        单词表内对应单词计数+1
        
        如果没有对应单词则插入该单词，初始化计数为1。并更新总单词量/出现次数。然后打印出当前操作记录-次数-频率-操作状态。
        注意！频率修改是局部的
        
        Args：
            word：str，需要更新的单词
        
        Return：
            stats：int，0失败，1更新，2插入
        """
        
        self.df.loc['_time_', "time"] += 1
        if word in self.df.index:  self.df.loc[word, "time"] += 1
        else:  
            self.df.loc[word] = [1, 0]
            self.df.loc['_word_', "time"] += 1
        
        self.update_frequency(word)
        print("Set word/time:  {}/{}".format(word, self.df.loc[word, "time"]))
    
    def get(self, word):
        """
        查询单词表，如果存在返回该单词以及对应数据，否则提示并返回error
        """
        if word in self.df.index: 
            self.update_frequency(word)
            print("Word:  ", word)
            print(self.df.loc[word])
        else:  print("Error, no this word: ", word)
    
    def update_frequency(self, word):
        self.df.loc[word, "frequency"] = self.df.loc[word, "time"]/self.df.loc['_time_', "time"]*100
    
    def update_all_frequency(self):
        """更新全部单词的频率"""
        for word in self.df.index:
            if "_" in word:  continue
            self.update_frequency(word)
    
    def show(self):
        """
        打印出整张表
        """
        self.update_all_frequency()
        print(self.df)
    
    def del_word(self, word):
        if word in self.df.index: 
            self.df.loc['_time_', "time"] -= self.df.loc[word, "time"]
            self.df.loc['_word_', "time"] -= 1
            self.df.drop(index=word, inplace=True)
            print("Del Word:  ", word)
        else:  print("Error, no this word: ", word)
    
    def save(self):
        self.update_all_frequency()
        self.df.to_excel(io)

    def lead_in(self, wordlist_path):
        """
        导入已有单词，以回车分隔的文本文档

        :param wordlist_path: 导入的文本路径
        """
        with open(wordlist_path, "r") as f:
            word_list = [i.rstrip("\n") for i in f.readlines()]
        print("Reading word file")
        for word in word_list:  self.set(word)