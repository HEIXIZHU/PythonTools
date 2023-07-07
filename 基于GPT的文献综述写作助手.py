#!/usr/bin/env python
# coding: utf-8

# In[ ]:


#方案名称：基于GPT的文献综述写作助手
背景：开题难，传统量化模型重复，高质量文献少、文献综述繁杂。
好的开题过程：先选定话题，再找出问题，然后明确对象，接着确定变量，最后套用理论按照理论的图式构建模型。
目标功能：
一、根据给定的话题，生成符合格式的参考题目；
二、批量爬取文献，收集理论模型，建立模型库；
三、根据题目，筛选并生成符合条件的研究方法、量化模型、文章结构；
四、根据爬取的文献进行自动的摘要、文章框架、文献综述写作。

挑战与解决方案：
1.设计知网爬虫与谷歌学术爬虫，还要正确清洗分类数据； → 用列表精选优质期刊、根据分类过滤无关文章
2.设计正确的Prompt，让GPT可以给出正确的回答； → 不断测试，设计正确的组合
3.设计正确的提问节奏，以免超出GPT的token限制。→ 将问题进行切割，每次返回的结果存储到变量中，最后再进行合并。


# In[ ]:


#文献综述自动撰写
# ✔可以加一个随机函数，将死板的“认为”偶尔改成“则指出”:效果不算好，尝试在prompt里用“以指出、基于开头”代替
# ✔加一个函数，清洗掉；号等多余的符号，名字间的则换成顿号
#加一个清洗函数，将“本文”、“这”、“本研究”等等不正确的写法删除替换。尝试在提问中加入限定要求，但效果不理想。
#检查是否有大量英文，如果有大量英文说明回传过程中出现了问题，要进行更改翻译。
#字数要有限制，过多的字数要求回传再次精简。
import pandas as pd
import xlrd
import time
import random
#调用ChatGPT
import openai
api_key = input('请输入API key：')
openai.api_key = api_key
def askChatGPT(messages):
    MODEL = "gpt-3.5-turbo"
    response = openai.ChatCompletion.create(
        model=MODEL,
        messages = messages,
        temperature=1)
    return response['choices'][0]['message']['content']
def SummaryAbbr(question):
        text = '下面是一篇论文的摘要，不允许出现“本文”、“本研究”、“论文”或“作者”、“我们”等词语和人称代词，如果有请删掉。按照前面的要求，对这篇摘要进行缩写,仅提取跟结论有关的内容：'+ question
        d = [{"role":"user","content":text}]
        text = askChatGPT(d)
        return text
def Wordsdown():
    pass

import os
path = input('输入文件夹路径(结尾注意添加斜杠）：')
#files= os.listdir(path)
#print(files)
files = ['UTAUT广告接受度.xlsx']
for filename in files:
    # 读取Excel文件
    df = pd.read_excel(filename)
    newlist = []
    # 遍历DataFrame中的每一行
    for index, row in df.iterrows():
        # 分别获取索引为'A','B','C'的单元格的值
        try:
            Sums = row['Summary-摘要']
            newSums = SummaryAbbr(Sums)
            while newSums == 'error':
                print(newSums)
                aaa = input('请确认网络正常后按ENTER')
                newSums = SummaryAbbr(Sums)
            else:
                newSums = SummaryAbbr(Sums)
            writetime = row['PubTime-发表时间']
            Authors = row['Author-作者']
            xxx = Authors.count(';')
            if xxx > 1:
                Authors = Authors.replace(';','、',xxx-1)
                Authors = Authors.replace(';','')
            else:
                Authors =Authors.replace(';','')
            point = ['认为','指出，','发现']    
            text = Authors + '（'+writetime[0:4]+'）'+point[random.randint(0,1)]+newSums
            print(text)
            newlist.append(text)
            print('timelimit')
            time.sleep(22)
            print('开始下一篇文献…')
        except:
            print('超出限额，请再等20秒…')
            time.sleep(22)
    with open (f'{filename}.txt','w') as f:
        f.writelines(newlist)
print('全部文献处理完毕。')
        

    



# In[ ]:


#调用chatGPT
import openai
api_key = input('请输入API key：')
openai.api_key = api_key
def askChatGPT(messages):
    MODEL = "gpt-3.5-turbo"
    response = openai.ChatCompletion.create(
        model=MODEL,
        messages = messages,
        temperature=1)
    return response['choices'][0]['message']['content']

def main():
    messages = [{"role": "user","content":""}]
    while 1:
        try:
            text = input('你：')
            if text == 'quit':
                break

            d = {"role":"user","content":text}
            messages.append(d)

            text = askChatGPT(messages)
            d = {"role":"assistant","content":text}
            print('ChatGPT：'+text+'\n')
            messages.append(d)
        except:
            messages.pop()
            print('ChatGPT：error\n')
main()

