

import os
import fnmatch
from win32com import client as wc


'''
Description: pdf 文件转txt文本
Author： prepared
Prompt：code in python3.6
'''
def file2Txt(filepath, savepath=''):
    try:
        print("")
        # 1 切分文件上级目录和文件名
        dirs, filename = os.path.split(filepath)

        # 2 修改文件后缀
        typename = os.path.splitext(filename)[-1]
        new_name = tranTye(filename, typename)

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


    except Exception as e:
        print(e)



'''
修改文件后缀
'''
def tranTye(filename, typename):
    new_name = ""
    typename = typename.lower()
    if typename == ".pdf":
        if fnmatch.fnmatch(filename, "*.pdf") or fnmatch.fnmatch(filename, "*.PDF"):
            new_name = filename[:-4] + ".txt"
        else:
            return
    elif typename == ".doc" or typename == '.docx':

        if fnmatch.fnmatch(filename, "*.doc"):
            new_name = filename[:-4] + ".txt"
        elif fnmatch.fnmatch(filename, "*.docx"):
            new_name = filename[:-5] + ".txt"
        else:
            return
    else:
        print("格式不正确，仅仅支持doc/docx/pdf")
    return new_name



if __name__ == "__main__":
    filepath = os.path.abspath(r'E:\work\ehualu\2018年中国城市智能交通市场研究报告.pdf')
    file2Txt(filepath)
    filepath2 = os.path.abspath(r'E:\work\05_重点对象监控\3_易慧分布式大数据存储计算框架(STDP-BP)-部署安装手册v1.0.doc')
    file2Txt(filepath2)


