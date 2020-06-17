#coding: utf-8
import httplib2 #os
from googleapiclient import discovery
from oauth2client import client, tools
from oauth2client.file import Storage
# from googleapiclient import errors
import base64,os
from pprint import pprint

# Gmail権限のスコープを指定
SCOPES = 'https://mail.google.com/'
# 認証ファイル
CLIENT_SECRET_FILE = 'credentials.json'
USER_SECRET_FILE = 'credentials-gmail.json'
# ------------------------------------
# ユーザ認証データの取得
def gmail_user_auth():
    store = Storage(USER_SECRET_FILE)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = 'Python Gmail API'
        credentials = tools.run_flow(flow, store, None)
        print('認証結果を保存しました:' + USER_SECRET_FILE)
    return credentials
# Gmailのサービスを取得
def gmail_get_service():
    credentials = gmail_user_auth()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('gmail', 'v1', http=http)
    return service
# ------------------------------------
# GmailのAPIが使えるようにする
service = gmail_get_service()

# メッセージを扱うAPI
messages = service.users().messages()
# 自分のメッセージ一覧を"maxResults"件得る
number_of_requests = int(input("取得したいメールの件数>"))
msg_list = messages.list(userId='me', maxResults=number_of_requests).execute()
pprint(msg_list)
#上記のメッセージID一覧を取得（idとthreadIdをリスト内の辞書として）
msg_ID_list = msg_list['messages']
pprint(msg_ID_list)

def get_dir_name():
    """同フォルダ内のフォルダ名をリストで取得"""
    path = "./"
    files = os.listdir(path)
    dir_name = [f for f in files if os.path.isdir(os.path.join(path, f))]
    print(dir_name)
    return dir_name
get_dir_name()


def LabelsDict(service, user_id):
    """ユーザーのメールボックスから対象となるlabelIDとラベル名の辞書を作成する"""
    """{'labelID':'ラベル名'}"""
    labels_dict = {}
    try:
        label_objects = service.users().labels().list(userId=user_id).execute()
        labels = label_objects['labels']
        print("以下labels")
        pprint(labels)
        """対象となるlabelの選定"""
        for label in labels:
            if 'インバータ（csvファイル）/' in label['name']:
                labels_dict[label['id'].replace("Label_","") ] = label['name'].replace('インバータ（csvファイル）/','')
        pprint(labels_dict)
        return labels_dict
    except Exception as e: 
        print("labels_dict取得する上で『{}』のエラー!".format(e))
        pass
LabelsDict(service,'me')

def GetAttachments(msg_ID_list):
    """添付ファイルを取得"""
    for msg in msg_ID_list:
        #メッセージIDを取得
        id = msg['id']
        # メッセージの本体を取得する
        data = messages.get(userId='me', id=id).execute()
        print("以下data")
        pprint(data)
        
        try:
            if data['payload']['parts'][1]['body']['attachmentId']:
                #添付ファイルが存在するファイルに対し、attachmentIDを取得
                attachment_ID = data['payload']['parts'][1]['body']['attachmentId']
                # 添付ファイルの本体を取得
                attachment = messages.attachments().get(userId='me',messageId = id, id=attachment_ID).execute()
                # 添付ファイルのコードを変換
                file_data = base64.urlsafe_b64decode(attachment['data'])
                #path名＝添付ファイル名とする
                print(f"添付ファイル名：{data['payload']['parts'][1]['filename']}")
                path = data['payload']['parts'][1]['filename']
                # open(path,"wb") as f: のpathはstrで./内にそのファイルが無くてもpath名で新規作成される。
                #https://note.nkmk.me/python-file-io-open-with/ ←参考
                with open(path,"wb") as f:
                    f.write(file_data)
        except KeyError: 
            continue
GetAttachments(msg_ID_list)

def LabelId_AttachmentId(msg_ID_list):
    """メッセージのlabelID、attachmentIdを取得"""
    """{'label_ID':'attachment_ID'}"""
    labelId_attachmentId_dict = {}
    for msg in msg_ID_list:
        #メッセージIDを取得
        id = msg['id']
        # メッセージの本体を取得する
        data = messages.get(userId='me', id=id).execute()
        pprint(data)
        try:
            #メッセージ本体に"labalIds"キーが存在した場合、
            if data['labelIds']:
                #"labelIDs"が存在するファイルに対し、labelIDを取得
                label_ID =  [i.replace("Label_","") for i in data['labelIds'] if "Label_" in i][0]
                print("右がlabel_ID→",label_ID)
                # attachmentIdを取得
                attachment_ID = data['payload']['parts'][1]['body']['attachmentId']
                # labelId_attachmentId_dictに上記二つを格納
                labelId_attachmentId_dict[label_ID] = attachment_ID
        except Exception as e: 
            print("labelId_attachmentId_dict取得する上で『{}』のエラー!".format(e))
            pass
    pprint(labelId_attachmentId_dict)
    return labelId_attachmentId_dict
LabelId_AttachmentId(msg_ID_list)

