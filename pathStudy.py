#!/usr/bin/python3

import os

# 绝对路径
path = os.getcwd()
print("path:" + path)

# 遍历当前路径下文件
files=os.listdir(path)
for file in files:
 print("file: "+file)

# 将相对路径转换为绝对路径
# path = "Excel"
# os.path.abspath(path)
# print(path)

# 将返回从 start 路径到 path 的相对路径的字符串。如果没有提供 start，就使用当前工作目录作为开始路径
# os.path.relpath(path)
# print(path)

# 判断是否是绝对路径
result = os.path.isabs(path)
print(result)


# 获取上一级目录
Path=os.path.dirname(os.getcwd())
print(Path)
# 获取上一级目录
Path=os.path.abspath(os.path.join(os.getcwd(),".."))
print(Path)