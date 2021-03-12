# -*- coding: UTF-8 -*-
import requests as req
import json,sys,time
import random,os
#先注册azure应用,确保应用有以下权限:
#files: Files.Read.All、Files.ReadWrite.All、Sites.Read.All、Sites.ReadWrite.All
#user:  User.Read.All、User.ReadWrite.All、Directory.Read.All、Directory.ReadWrite.All
#mail:  Mail.Read、Mail.ReadWrite、MailboxSettings.Read、MailboxSettings.ReadWrite
#注册后一定要再点代表xxx授予管理员同意,否则outlook api无法调用

path=sys.path[0]+r'/1.txt'
num1 = 0

# refresh_token = os.environ.get('REFRESH_TOKEN')
id = os.environ.get('CLIENT_ID')
secret = os.environ.get('CLIENT_SECRET')

def gettoken(refresh_token):
    headers={'Content-Type':'application/x-www-form-urlencoded'
            }
    data={'grant_type': 'refresh_token',
          'refresh_token': refresh_token,
          'client_id':id,
          'client_secret':secret,
          'redirect_uri':'http://localhost:53682/'
         }
    html = req.post('https://login.microsoftonline.com/common/oauth2/v2.0/token',data=data,headers=headers)
    jsontxt = json.loads(html.text)
    try:
        refresh_token = jsontxt['refresh_token']
    except KeyError:
        print(jsontxt)
    access_token = jsontxt['access_token']
    with open(path, 'w+') as f:
        f.write(refresh_token)
    return access_token

all_API_url = [r'https://graph.microsoft.com/v1.0/me/drive/root',
r'https://graph.microsoft.com/v1.0/me/drive',
r'https://graph.microsoft.com/v1.0/drive/root',
r'https://graph.microsoft.com/v1.0/users',
r'https://graph.microsoft.com/v1.0/me/messages',
r'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules',
r'https://graph.microsoft.com/v1.0/me/mailFolders/Inbox/messages/delta',
r'https://graph.microsoft.com/v1.0/me/drive/root/children',
r'https://api.powerbi.com/v1.0/myorg/apps',
r'https://graph.microsoft.com/v1.0/me/mailFolders',
r'https://graph.microsoft.com/v1.0/me/outlook/masterCategories']

num1 = 0
with open(path, "r+") as fo:
    refresh_token = fo.read()

def main():
    # localtime = time.asctime( time.localtime(time.time()) )
    access_token=gettoken(refresh_token)
    headers={
    'Authorization':access_token,
    'Content-Type':'application/json'
    }
    begin = random.randint(0,len(all_API_url))
    end = random.randint(begin,len(all_API_url))
    for x in range(begin,end):
        try:
            r = req.get(all_API_url[x],headers=headers)
            if r.status_code == 200:
                num1+=1
                print(f'{x}号发射成功, 总第{num1}次成功')
            else:
                print(r.text)
        except:
            print(f"{x}号发射失败")
            pass
        if x == end-1 :
            print('此次运行结束时间为 :', time.asctime( time.localtime(time.time()) ) )
for i in range(3):
    main()
