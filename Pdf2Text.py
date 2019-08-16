#coding=utf-8

"""
Description: pdf 文件转txt文本
Author： prepared
Prompt：code in python3.6
"""

import os
import fnmatch
from win32com import client as wc


def pdf2Txt(filepath, savepath=''):
    # 1 切分文件上级目录和文件名
    dirs, filename = os.path.split(filepath)
    print("filename>>>>"+filename)

    # 2 修改文件后缀
    new_name = ""
    if fnmatch.fnmatch(filename, "*.pdf") or fnmatch.fnmatch(filename, "*.PDF"):
        new_name = filename[:-4] + ".txt"
    else:
        print("文件格式不正确，只支持pdf")
        return
    print("new_name>>>>"+new_name)

    # 3 获取保存路径
    if savepath == '':
        savepath = dirs
    else:
        savepath = savepath
    pdf_to_txt = os.path.join(savepath, new_name)
    print("pdf_to_txt>>>>"+pdf_to_txt)

    # pdfapp = wc.Dispatch('Word.Application')
    # mytxt = pdfapp.Documents.Open(filepath)
    # mytxt.SaveAs(pdf_to_txt, 4)
    # print(mytxt)
    # mytxt.Close()
    wordapp = wc.Dispatch('Word.Application')
    mytxt = wordapp.Documents.Open(filepath)
    print("filepath>>>>"+filepath)
    print("mytxt>>>>"+mytxt)

    mytxt.SaveAs(pdf_to_txt, 4)
    mytxt.Close()


if __name__ == "__main__":
    filepath = os.path.abspath(r'E:\work\ehualu\2018年中国城市智能交通市场研究报告.pdf')
    pdf2Txt(filepath)









