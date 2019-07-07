#!/bin/python
import os
import sys

def hmOpenPython3(fileName ,flag):
    r = os.system("python E:\\hmbi\\hmPY\\hmPdf0706.cpython-37.pyc %s:%s"% (fileName,flag))
    return r

if __name__ == '__main__':
    shellValues = sys.argv[1]
    shellArr =  shellValues.split(",")
    flag = int(shellArr[1])
    fileName = shellArr[0]
    res = hmOpenPython3(fileName,flag)
