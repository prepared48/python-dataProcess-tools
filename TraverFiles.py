# coding=utf-8

import  os,time


class TraversalFun():

    # 1 初始化
    def __init__(self, rootDir):
        self.rootDir = rootDir


    def TraversalDir(self):
        TraversalFun.AllFiles(self, self.rootDir)


    def AllFiles(self, rootDir):
        for file in os.listdir(rootDir):
            path  = os.path.join(rootDir, file)
            if os.path.isfile(path):
                print(os.path.abspath(path))
            elif os.path.isdir(path):
                TraversalFun.AllFiles(self, path)


if __name__ == "__main__":
    startTime = time.time()
    tra = TraversalFun("E:/prepared/100-Days-Of-ML-Code-master")

    tra.TraversalDir()
    endTime = time.time()
    print("total cost time(s)", (endTime-startTime), 's')