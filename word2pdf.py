#!/usr/bin/env python
#created by Maosong Pei

#python模块:win32com用法:
#import win32com
#from win32com.client import Dispatch, constants
#w = win32com.client.Dispatch('Word.Application')
# 打开新的文件
#doc = w.Documents.Open( FileName = filenamein )
#walk用法
#for root, dirs, filenames in walk(directory):
#root 所指的是当前正在遍历的这个文件夹的本身的地址
#dirs 是一个 list ，内容是该文件夹中所有的目录的名字(不包括子目录)
#files 同样是 list , 内容是该文件夹中所有的文件(不包括子目录)



from win32com.client import Dispatch
from os import walk

wdFormatPDF = 17


def doc2pdf(input_file):
    word = Dispatch('Word.Application')
    doc = word.Documents.Open(input_file)
    doc.SaveAs(input_file.replace(".docx", ".pdf"), FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()


if __name__ == "__main__":
    doc_files = []
    path=os.getcwd()
    directory = path
    for root, dirs, filenames in walk(directory):
        for file in filenames:
            if file.endswith(".doc") or file.endswith(".docx"):
                doc2pdf(str(root + "\\" + file))
