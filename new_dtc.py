import requests
import pandas
import openai
import time
import datetime

from urllib.parse import urlparse, parse_qs
APP_ID='cli_a6109595397dd00c'
APP_SECRET='4kvnWi7eSELx5lLyQv7NkgsRJMoBL8yc'
REDIRECT_URI = 'http://123.56.166.61/redirect'  # 云服务器的重定向地址
SRC_SPREADSHEET_TOKEN='RbqqsliPPhYvRUtQgeCcZ5hxnzg'
SRC_SHEET_TITLE='Sheet1'
DST_SHEET_TITLE='Sheet2'
USER_ID='g68857da'
OPENAI_API_BASE='http://localhost:3000'
OPENAI_API_KEY='sk-J5fhnUmY1GeAxq5H3c3102F815174f639b95A012816a526e'
#自定义开始行数
MY_BEGIN=1727
MY_END=1731


#OPENAI_API_BASE='https://aiserver.marsyoo.com/'
#OPENAI_API_KEY='sk-uB4cBeoA8HZLOQ9K62965eD12531414d9d4fDe04Ca81418'



#成功时返回需要的报文，失败时返回-1
#获取user_access_token
def get_access_token(app_id, app_secret):
    url = 'https://open.feishu.cn/open-apis/auth/v3/app_access_token/internal'
    headers = {'Content-Type': 'application/json'}
    payload = {
        'app_id': app_id,
        'app_secret': app_secret
    }
    response = requests.post(url, json=payload, headers=headers)#
    response_data = response.json()
    if response_data['code']!=0:
        return -1
    return response_data['app_access_token']
def get_authorization_url(app_id, redirect_uri):
    url = f"https://open.feishu.cn/open-apis/authen/v1/index?app_id={app_id}&redirect_uri={redirect_uri}&response_type=code&scope=snsapi_userinfo"
    return url

def extract_pre_auth_code(redirected_url):#拿到登陆预授权码
    parsed_url = urlparse(redirected_url)
    query_params = parse_qs(parsed_url.query)
    pre_auth_code = query_params.get('code', [None])[0]
    return pre_auth_code

def get_user_access_token(app_access_token,pre_auth_code):
    url='https://open.feishu.cn/open-apis/authen/v1/access_token'
    headers = {'Authorization':f'Bearer {app_access_token}',
               'Content-Type': 'application/json'}
    payload = {
        'grant_type': 'authorization_code',
        'code': pre_auth_code
    }
    response = requests.post(url, json=payload, headers=headers)  #
    response_data=response.json()
    if response_data['code']!=0:
        return -1
    return response_data['data']['access_token']
##################################################################
#读取[begin行,end行]这个区间内的信息,其中begin从1开始
def get_src_sheet_rows(spreadsheet_token, user_access_token):
    #先想办法提取src_sheet一共有多少行
    url = f"https://open.feishu.cn/open-apis/sheets/v3/spreadsheets/{spreadsheet_token}/sheets/query"
    headers = {
        "Authorization": f"Bearer {user_access_token}",
    }
    rsp = requests.get(url, headers=headers)
    if rsp.json()['code'] != 0:
        print("获取rows失败\n")
        return -1
    sheets_val = rsp.json()['data']['sheets']
    #list = []
    if sheets_val[0]['title'] == SRC_SHEET_TITLE:
        rows=sheets_val[0]['grid_properties']['row_count']
        #list.append(sheets_val[0]['sheet_id'])
        #list.append(sheets_val[1]['sheet_id'])
    else:
        rows = sheets_val[1]['grid_properties']['row_count']
        #list.append(sheets_val[1]['sheet_id'])
        #list.append(sheets_val[0]['sheet_id'])

    return rows

#把读取出来的报文组织成想要的，能直接向chatgpt提问的格式
def organize_data(report):
    organized_data = []
    value_range = report.get('valueRange', {}).get('values', [])

    for entry in value_range:
        for content in entry:
            if isinstance(content, list):
                for item in content:
                    if item.get('type') == 'mention':
                        link = item.get('link', '')
                        text = item.get('text', '')
                        organized_data.append([link, text, '', ''])

    return organized_data
#generate_summary的返回值是可以直接拿出来插入到飞书中的内容
def generate_summary(data):
    #答案是根本不存在api_base这个成员,必须自己重头组织url
    url=f"{OPENAI_API_BASE}/v1/chat/completions"

    #初始化对对OpenAI模型的指示，告诉模型将要处理的内容和任务。这段指示文本是一个引导，帮助模型理解接下来的内容
    prompt = "以下是一些文章的链接和对应标题，请根据链接总结出每个文章的关键词(2个)、摘要(不少于30字)和更新日志,并按照形如[['链接1','关键字a,关键字b','摘要1','更新日志1'],['链接2','关键字c,关键字d','摘要2','更新日志2']]的顺序和格式回答,无有效内容时回答[]\n\n"
    # 遍历 data 列表中的每个 item,将每个 item 中的链接和文本信息追加到 prompt 字符串中
    for item in data:
        prompt += f"链接: {item[0]}\n内容: {item[1]}\n\n"
    headers = {
        'Authorization': f'Bearer {OPENAI_API_KEY}',
        'Content-Type':'application/json'
    }
    payload = {
        "model": "gpt-4o",
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.7
    }
    response = requests.post(url, json=payload, headers=headers)
    rsp = response.json()
    content=rsp['choices'][0]['message']['content']

    try:
        content_list=eval(content)
        return content_list
    except (SyntaxError,KeyError):
        #对应一点信息也提取不出来的情况，对应非法区间状态
        return -314


#直接复用了以上三个函数
#使用的接口本质上是对已有表格进行更新，所以必须保证dst_sheet有大于等于src_sheet的行数
def insert_src_sheet(spreadsheet_token, user_access_token,begin,end):
    # 先想办法提取src_sheet一共有多少行
    flag=0
    url = f"https://open.feishu.cn/open-apis/sheets/v3/spreadsheets/{spreadsheet_token}/sheets/query"
    headers = {
        "Authorization": f"Bearer {user_access_token}",
    }
    rsp = requests.get(url, headers=headers)
    if rsp.json()['code'] != 0:
        print("获取rows失败\n")
        return -1
    sheets_val = rsp.json()['data']['sheets']
    list = []
    if sheets_val[0]['title'] == SRC_SHEET_TITLE:
        rows = sheets_val[0]['grid_properties']['row_count']
        list.append(sheets_val[0]['sheet_id'])
        list.append(sheets_val[1]['sheet_id'])
    else:
        rows = sheets_val[1]['grid_properties']['row_count']
        list.append(sheets_val[1]['sheet_id'])
        list.append(sheets_val[0]['sheet_id'])
    # 注意拿到的rows就是行数
    # 开始更新闭区间内的数据
    if begin < 2:
        print("begin非法\n")
        return -1

    #rows很可能是在插入表格时就确定的，清数据不改变rows但是直接用删除选项会改变rows
    #能解释源spreadsheet和副本行数不同的现象
    #结论是rows并不可靠,不如直接看元素
    #################################################
    q = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{spreadsheet_token}/values/{list[0]}!B{begin}:B{end}"
    headers = {
        "Authorization": f"Bearer {user_access_token}",
        "Content-Type": "application/json; charset=utf-8"
    }
    w = requests.get(q, headers=headers)
    e = w.json()
    if e['code'] != 0:
        print("读取表格内容失败\n")
        print(e['msg'])
        return -1
    raw_data= e['data']

    ogn_data=organize_data(raw_data)
    content=[]
    content=generate_summary(ogn_data)
    if content==-314:
        return 0
    #len为返回content里面元素的个数,神来之笔,rows不可靠所以必须以csz来判断有效数据边界
    csz=len(content)
    #[1727,1731]返回5个元素
    end = begin + csz - 1
    if(csz==0):
        return 0
    ######################################################################
    # 到这里拿到了两个工作表的id:list[],srcsheet行数rows
    # 以及openai组织好的可以直接插入表格的content
    #print(content)#success
    ##############插入数据以后的部分：
    url=f'https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{spreadsheet_token}/values'
    headers={
        'Authorization': f'Bearer {user_access_token}',
        'Content-Type': 'application/json; charset=utf-8'
    }

    payload={
        'valueRange':{
            'range':f'{list[1]}!A{begin}:E{end}',
            'values':content
        }
    }
    rsp=requests.put(url, json=payload, headers=headers)
    if rsp.json()['code'] == 0 :
        print(f'{begin}-{end}行写入成功\n')
        #返回成功写入的行数
        return csz

    else:
        print(rsp.json()['msg'])
        #print(rsp)
        return -1
###############################################################################3
#模块直接把src_sheet里的所有内容全部导入到新表格
def moudle_start():
    # 获取 app_access_token
    access_token = get_access_token(APP_ID, APP_SECRET)  # app_access_token
    # 提取预授权码
    authorization_url = get_authorization_url(APP_ID, REDIRECT_URI)
    print(f"请访问以下链接并授权：\n{authorization_url}")
    redirected_url = input("请将授权后重定向的URL粘贴到这里：\n")
    pre_auth_code = extract_pre_auth_code(redirected_url)
    # 获取user_access_token
    user_access_token = get_user_access_token(access_token, pre_auth_code)
    begin = MY_BEGIN
    end = MY_END
    finished=begin-1
    while 1:
        rows = get_src_sheet_rows(SRC_SPREADSHEET_TOKEN, user_access_token)
        #insert_src_sheet函数中设置了对[begin,end]区间的调整
        while begin<=rows:
            #k为从开始的位置开始读了多少行
            k = insert_src_sheet(SRC_SPREADSHEET_TOKEN, user_access_token, begin, end)
            if k==0:
                break
            elif k>0:#正在更新的区间[begin,end]中包含边界
                #[1327,1331]返回的是5
                finished=begin+k-1
                begin=finished+1
                if k==5:
                    end=begin+5-1
                elif k>0&k<5:
                    #到边界了,不能再读了
                    end=begin

            #发现一旦k!=0,数据更新就会停止
            else:
                print(f'[{begin},{end}]区间插入数据失败,正在重新获取user_access_token,请重新授权\n')
                access_token = get_access_token(APP_ID, APP_SECRET)  # app_access_token
                # 提取预授权码
                authorization_url = get_authorization_url(APP_ID, REDIRECT_URI)
                print(f"请访问以下链接并授权：\n{authorization_url}")
                redirected_url = input("请将授权后重定向的URL粘贴到这里：\n")
                pre_auth_code = extract_pre_auth_code(redirected_url)
                # 获取user_access_token
                user_access_token = get_user_access_token(access_token, pre_auth_code)
        #每10s检查一次表格是否需要再次更新
        #time.sleep(10)#测试环境下无误
        time.sleep(600)
        print('工作表内容更新完毕')



if __name__ == "__main__":
    moudle_start()




















