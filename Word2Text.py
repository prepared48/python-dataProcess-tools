#coding=utf-8

"""
Description: word文件转txt文本
Author： prepared
Prompt：code in python3.6
"""

import os
import fnmatch
from win32com import client as wc


def word2Txt(filepath, savepath=''):
    # 1 切分文件上级目录和文件名
    dirs, filename = os.path.split(filepath)

    # 2 修改文件后缀
    new_name = ""
    if fnmatch.fnmatch(filename, "*.doc"):
        new_name = filename[:-4] + ".txt"
    elif fnmatch.fnmatch(filename, "*.docx"):
        new_name = filename[:-5] + ".txt"
    else:
        print("文件格式不正确，只支持doc和docx")
        return

    # 3 获取保存路径
    if savepath == '':
        savepath = dirs
    else:
        savepath = savepath
    word_to_txt = os.path.join(savepath, new_name)

    wordapp = wc.Dispatch('Word.Application')
    mytxt = wordapp.Documents.Open(filepath)
    mytxt.SaveAs(word_to_txt, 4)
    print(mytxt)
    mytxt.Close()


if __name__ == "__main__":
    filepath = os.path.abspath(r'E:\work\05_重点对象监控\3_易慧分布式大数据存储计算框架(STDP-BP)-部署安装手册v1.0.doc')
    word2Txt(filepath)






