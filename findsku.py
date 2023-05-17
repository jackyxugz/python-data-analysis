# encoding:utf-8
import sys
import  os
import requests
# import urllib
# import urllib.request
import json
import time
import jieba
from gensim import corpora,models,similarities
import pandas as pd
# from tkinter import filedialog

# client_id 为官网获取的AK， client_secret 为官网获取的SK
# host = 'https://aip.baidubce.com/oauth/2.0/token?grant_type=client_credentials&client_id=Ntr9iYThVm69GSt0RPAmDH5K&client_secret=KZiVTPoqzirPgdZk11jVc9ENIWIMx3Nw'
# response = requests.get(host)
# if response:
#     print(response.json())


'''
  var APP_ID = "10661815";
            var API_KEY = "Ntr9iYThVm69GSt0RPAmDH5K";
            var SECRET_KEY = "KZiVTPoqzirPgdZk11jVc9ENIWIMx3Nw";
'''

# 设置文件对话框会显示的文件类型
my_filetypes = [('text excel files', '.xlsx'),('all excel files', '.xls')]
# title粗略匹配纪录数限制
limit1=20
# title精准匹配纪录数限制
limit2=20
# 补充上编码以后的记录数限制
limit3=100


class BaiduNLP:

    def __init__(self, client_id, client_secret):
        self.session = requests.Session()
        self.client_id = client_id
        self.client_secret = client_secret

    def getAccessToken(self):
        params = {
            'grant_type': 'client_credentials',
            'client_id': self.client_id,
            'client_secret': self.client_secret,
        }
        response = self.session.post(
            'https://aip.baidubce.com/oauth/2.0/token', params=params)
        jsonres = json.loads(response.content)
        # print(jsonres)
        return jsonres.get('access_token')

    def getResult(self, host, text_1,text_2, options):
        access_token = "24.1984faa19ddd0d2a5ddb1c3b7e3d9557.2592000.1644134150.282335-10661815"
            # self.getAccessToken()
        # print(access_token)
        myhost = "%s?access_token=%s&charset=UTF-8" % (host, access_token)
        headers = {'Content-Type': 'application/json'}
        # params['charset'] = 'UTF-8'
        # data = {'text_1': text_1,'text_2': text_2}
        # data = '{"text_1": "星空漱口水", "text_2": "漱口水"}'
        # data = '{"text_1": "'+text_1+'", "text_2": "'+text_2+'"}'
        data = '{"text_1": "%s", "text_2": "%s"}' % (text_1, text_2)
        # data = '{"text_1": "{}", "text_2": "{}"}'.format(,text_2)
        # print("data=")
        # print(data)
        data=data.encode("UTF-8")
        # return ""
        # data = '{"text_1": "\u6d59\u5bcc\u80a1\u4efd", "text_2": "\u4e07\u4e8b\u901a\u81ea\u8003\u7f51"}'
        # data.update(options)
        # data = json.dumps(data)
        response = self.session.post(myhost, headers=headers, data=data)
        return response.content.decode()


def cal_xiangsidu_baidu(str1,str2):

    APP_ID = '10661815'
    API_KEY = 'Ntr9iYThVm69GSt0RPAmDH5K'
    SECRET_KEY = 'KZiVTPoqzirPgdZk11jVc9ENIWIMx3Nw'

    host = "https://aip.baidubce.com/rpc/2.0/nlp/v2/simnet"
    options = {
    }

    client = BaiduNLP(API_KEY, SECRET_KEY)
    result = client.getResult(host, str1,str2, options)
    # print(result)
    # json_str = json.dumps(result)

    # 将 JSON 对象转换为 Python 字典
    data2 = json.loads(result)
    # print("精确比对结果:")
    # print(data2)
    return  data2["score"]
    # return  data2[1]



def get_key1(title):
    # doc_test = "我喜欢上海的小吃"
    # print("title:")
    # print(title)
    doc_test ="".join(title)
    doc_test_list = [word for word in jieba.cut(doc_test)]
    return doc_test_list


def cal_xiangsidu(doc_test_list,all_doc_list,match_percent):

    # 用dictionary方法获取词袋（bag-of-words)
    dictionary = corpora.Dictionary(all_doc_list)
    # 词袋中用数字对所有词进行了编号
    # print(dictionary.keys())
    # 编号与词之间的对应关系
    # print(dictionary.token2id)
    # 使用doc2bow制作语料库
    corpus = [dictionary.doc2bow(doc) for doc in all_doc_list]

    # 用同样的方法，把测试文档也转换为二元组的向量
    doc_test_vec = dictionary.doc2bow(doc_test_list)
    # print(doc_test_vec)
    # 相似度分析，使用TF - IDF模型对语料库建模
    tfidf = models.TfidfModel(corpus)
    # 获取测试文档中，每个词的TF - IDF值
    # print(tfidf[doc_test_vec])
    # 对每个目标文档，分析测试文档的相似度
    index = similarities.SparseMatrixSimilarity(tfidf[corpus], num_features=len(dictionary.keys()))
    sim = index[tfidf[doc_test_vec]]
    # print("未排序")
    # print(sim)
    sim=sorted(enumerate(sim), key=lambda item: -item[1])
    # print("已排序")
    # print(sim)
    # print(sim[0][0])
    # print("".join(all_doc_list[sim[0][0]]),"".join(all_doc_list[sim[1][0]]),"".join(all_doc_list[sim[2][0]]))
    # print( sim[0][1],sim[1][1],sim[2][1] )

    # print("测试:")
    row=[]

    for i in range(0,limit1):
        if len(sim)>i:
            item="".join(all_doc_list[sim[i][0]])
            percent=sim[i][1]
            # 匹配精度
            if percent>=match_percent:
                row.append([item,percent])

    # df=pd.Dataframe(row).reset_index()
    # print("打印结果:", "".join(doc_test_list))
    # print(row)
    # print(df)

    return  row

    # all_doc_list

def read_vouch(filename):
    df=pd.read_excel(filename)

    return df


def read_product(filename):
    df = pd.read_excel(filename)
    df.columns=["key","title"]

    all_doc_list = []
    # for doc in all_doc:
    for index,doc in df.iterrows():
        doc_list = [word for word in jieba.cut(doc["title"])]
        all_doc_list.append(doc_list)

    # print(all_doc_list)
    return  all_doc_list


def get_percent(title,key2,gross_match_percent):
    # title = "苗坚防脱育发洗发水（矮瓶子）300ML塑料瓶"
    key1 = get_key1(title)

    gross_list = cal_xiangsidu(key1, key2,gross_match_percent)
    # print("粗略匹配:",key1)
    # print(gross_list)
    if len(gross_list)==0:
        gross_list = cal_xiangsidu(key1, key2, gross_match_percent/2)
        # print("粗略匹配2:", key1)
        # print(gross_list)

    if len(gross_list)==0:
        gross_list = cal_xiangsidu(key1, key2, gross_match_percent/4)
        # print("粗略匹配3:", key1)
        # print(gross_list)


    # print("精细匹配")
    # print(gross_list)

    result_list = []
    cnt=0
    for s in gross_list[0:limit1+1]:
        key2 = s[0]
        p=s[1]
        if p>0.95:
            # print("粗略匹配合格，不需要精准匹配")
            pass
        else:
            p = cal_xiangsidu_baidu(title, key2)
            # print("{} 的匹配精度是： {}".format(key2,p))
            # result_list.append((key2,p))
            cnt=cnt+1
        result_list.append([key2, p])

    # print("调用{}次百度AI".format(cnt))
    # print(result_list)
    # print("最后匹配的结果")
    result_list = sorted(result_list, key=lambda item: -item[1])
    # print(result_list)
    return result_list

def zhaodao(title,key2,gross_match_percent,baidu_match_percent):
    # key2 = read_product(filename)
    df_result=[]
    # 计算匹配度
    percent_list=get_percent(title,key2,gross_match_percent)
    # print("计算匹配度结果:")
    # print(percent_list)
    v_limit2=limit2+1
    for i in range(0,v_limit2):
        if len(percent_list)>i:  # 如果没有越界
            # print(title,"匹配匹配；",percent_list[i])
            df_result.append([title,percent_list[i][0],percent_list[i][1]])
            # print(df_result)

    df_result=pd.DataFrame(df_result).reset_index()
    if df_result.empty:
        print("没有匹配到记录")

    else:
        # print("粗略计算完以后的记录数:",df_result.shape[0])
        # print("测试")
        # print(df_result.head(10).to_markdown())
        df_result.columns=["index","title1","title2","percent"]
        # 恢复原始的产品名称
        df_result["title2"]=df_result["title2"].apply(lambda x:  x.split("|")[0] )
        # print("比对完的结果:")
        # print(df_result.to_markdown())
        if df_result[df_result["percent"]>=baidu_match_percent].shape[0]>0:
            df_result=df_result[df_result["percent"]>=baidu_match_percent]
        elif df_result[df_result["percent"]>=baidu_match_percent/2].shape[0]>0:
            df_result=df_result[df_result["percent"]>=baidu_match_percent/2]
        elif df_result[df_result["percent"] >= baidu_match_percent / 4].shape[0] > 0:
            df_result = df_result[df_result["percent"] >= baidu_match_percent / 4]

        # print("对比合格的记录数:",df_result.shape[0])
        # df_result.to_excel(r"C:\Users\ns2033\Downloads\bom名称比对结果4.xlsx")
        # print(df_result.to_markdown())

        # print("百度精准计算完以后的记录数:", df_result.shape[0])

    # v_limit2 = min(v_limit2, df_result.shape[0])
    # v_limit2=v_limit2 + 1
    return df_result.head(limit2)


if __name__ == "__main__":
    # print("开始:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
    # xsd=cal_xiangsidu("若蘅视黄醇紧致舒纹精华乳100ml真空瓶","若蘅视黄醇紧致 100ml真空瓶")
    # print(xsd)

    print("请输入词典文件路径:")

    # 请求选择文件
    # filename = filedialog.askopenfilename(initialdir=os.getcwd(),
    #                                       title="请输入词典文件路径",
    #                                       filetypes=my_filetypes)
    #
    # if len(filename) == 0:
    #     sys.exit()
    #
    # print("你选择的文件名是：", filename)
    # # return filename

    filename=input()
    # filename = sys.stdin.readline()
    filename=filename.strip()
    if len(filename)==0:
        print("你没有输入词典，程序退出！")
        # sys.exit()
    else:
        print("读取词典:",filename,"...")
        df = pd.read_excel(filename)
        df.columns = ["key", "title"]
        key2 = read_product(filename)

        while True:
        # for i in range(1,3):
            print("请输入要匹配的文本:")
            key = input()
            if len(key)>0:
                key=key.replace("；",";")
                if key.find(";")>0:
                    for i in range(0,len(key.split(";"))):
                        c_key=key.split(";")[i]
                        df1 = zhaodao(c_key, key2, 0.7, 0.1)
                        df2 = df1.merge(df, how="left", left_on=["title2"], right_on=["title"])
                        df3 = df2[["title1", "key", "title2", "percent"]]
                        # limit3=min(limit3,df3.shape[0])
                        # limit3=50
                        print(df3.head(limit3).to_markdown())
                else:
                    df1=zhaodao(key,key2,0.7,0.1)
                    df2=df1.merge(df,how="left",left_on=["title2"],right_on=["title"])
                    df3=df2[["title1","key","title2","percent"]]
                    # limit3=min(limit3,df3.shape[0])
                    # limit3=50
                    print(df3.head(limit3).to_markdown())
