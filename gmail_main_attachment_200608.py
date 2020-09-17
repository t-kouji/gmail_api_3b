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
# msg_list = messages.list(userId='me', maxResults=number_of_requests).execute()
# 上記のメッセージID一覧を取得（idとthreadIdをリスト内の辞書として）
# msg_ID_list = msg_list['messages']

def MessageDataDict(number_of_requests):
    """メッセージIDとメッセージ本体の辞書を作成する"""
    """{'id':'data'}"""
    message_data_dict = {}
    msg_list = messages.list(userId='me', maxResults=number_of_requests).execute()
    msg_ID_list = msg_list['messages']
    for msg in msg_ID_list:
        #メッセージIDを取得
        id = msg['id']
        # メッセージの本体を取得する
        data = messages.get(userId='me', id=id).execute()
        message_data_dict[id] = data
    return message_data_dict
MessageDataDict = MessageDataDict(number_of_requests)

def LabelsDict(service, user_id):
    """ユーザーのメールボックスから対象となるlabelIDとラベル名の辞書を作成する"""
    """{'labelID':'labelname'}"""
    labelsdict = {}
    try:
        label_objects = service.users().labels().list(userId=user_id).execute()
        labels = label_objects['labels']
        """対象となるlabelの選定"""
        for label in labels:
            if 'インバータ（csvファイル）/' in label['name']:
                labelsdict[label['id'].replace("Label_","") ] = label['name'].replace('インバータ（csvファイル）/','')
        return labelsdict
    except Exception as e: 
        print("labelsdict取得する上で『{}』のエラー!".format(e))
        pass
LabelsDict = LabelsDict(service,'me')
# pprint(LabelsDict)

def LabelId_AttachmentId(MessageDataDict):
    """メッセージのlabelID、attachmentIdを取得"""
    """{'label_ID':'attachment_ID'}"""
    """但し、最新のメールからforループを回すので、同じ案件でダブると最新のが上書きされて古いメールが
    残ってしまうのでそうならないようにする"""
    LabelId_AttachmentId_dict = {}
    for id,data in MessageDataDict.items():
        try:
            #メッセージ本体に"labalIds"キーが存在した場合、
            if data['labelIds']:
                #"labelIDs"が存在するファイルに対し、labelIDを取得
                label_ID =  [i.replace("Label_","") for i in data['labelIds'] if "Label_" in i][0]
                # attachmentIdを取得
                attachment_ID = data['payload']['parts'][1]['body']['attachmentId']
                # LabelId_AttachmentId_dictに上記二つを格納
                # 最新メールのみが反映されるようにif not inで条件付けする。
                if  label_ID not in LabelId_AttachmentId_dict:
                    LabelId_AttachmentId_dict[label_ID] = attachment_ID
        except Exception as e: 
            print("LabelId_AttachmentId_dict取得する上で『{}』のエラー!".format(e))
            pass
    return LabelId_AttachmentId_dict
LabelId_AttachmentId = LabelId_AttachmentId(MessageDataDict)
#pprint(LabelId_AttachmentId)

def Labelname_AttachmentId():
    """labelnameとattachmentIdの辞書を取得"""
    """{'Labelname':'attachment_ID'}"""
    Labelname_AttachmentId_dict = {}
    for label_ID,attachment_ID in LabelId_AttachmentId.items():
        for labelID,labelname in LabelsDict.items():
            if labelID in label_ID:
                Labelname_AttachmentId_dict[LabelsDict[labelID]] = LabelId_AttachmentId[label_ID]
    return Labelname_AttachmentId_dict
Labelname_AttachmentId = Labelname_AttachmentId()
#pprint(Labelname_AttachmentId)

# def Sort_Attachment(MessageDataDict):
#     """添付ファイルを各フォルダのdataフォルダ内に格納"""
#     for Labelname,attachment_ID in Labelname_AttachmentId.items():
#         """フォルダが存在しない場合は名前を指定してフォルダを作成。
#         第二引数にexist_ok=Trueとすると既に末端ディレクトリが存在している場合もエラーが発生しない"""
#         os.makedirs('./projects/{}/data'.format(Labelname),exist_ok=True)
#         """添付ファイル本体の取得"""
#         try:
#             for id,data in MessageDataDict.items():
#                 #該当するattachment_IDだった場合
#                 if data['payload']['parts'][1]['body']['attachmentId'] == attachment_ID:
#                     """添付ファイルの本体を取得"""
#                     attachment =  messages.attachments().get(userId='me',messageId = id, id=attachment_ID).execute()
#                     """添付ファイルコードの変更"""
#                     file_data = base64.urlsafe_b64decode(attachment['data'])
#                     """添付ファイル名の取得"""
#                     file_name = data['payload']['parts'][1]['filename']
#                     print(f"添付ファイル名：{file_name}")
#                     path = './projects/{}/data/{}'.format(Labelname,file_name) 
#                     with open(path,"wb") as f:
#                         f.write(file_data)
#                     # open(path,"wb") as f: のpathはstrで./内にそのファイルが無くてもpath名で新規作成される。
#                     # https://note.nkmk.me/python-file-io-open-with/ ←参考
#         except Exception as e: 
#             print("Sort_Attachmentで添付ファイルを振り分ける上で『{}』のエラー!".format(e))
#             continue
# Sort_Attachment(MessageDataDict)

def Sort_Attachment(MessageDataDict):
    """添付ファイルを各フォルダのdataフォルダ内に格納"""
    for Labelname,attachment_ID in Labelname_AttachmentId.items():
        """フォルダが存在しない場合は名前を指定してフォルダを作成。
        第二引数にexist_ok=Trueとすると既に末端ディレクトリが存在している場合もエラーが発生しない"""
        os.makedirs('./projects/{}/data'.format(Labelname),exist_ok=True)
        """添付ファイル本体の取得"""
        for id,data in MessageDataDict.items():
            # try:
            #該当するattachment_IDだった場合
            try:
                if data['payload']['parts'][1]['body']['attachmentId'] == attachment_ID:
                    """添付ファイルの本体を取得"""
                    attachment =  messages.attachments().get(userId='me',messageId = id, id=attachment_ID).execute()
                    """添付ファイルコードの変更"""
                    file_data = base64.urlsafe_b64decode(attachment['data'])
                    """添付ファイル名の取得"""
                    file_name = data['payload']['parts'][1]['filename']
                    print(f"添付ファイル名：{file_name}")
                    path = './projects/{}/data/{}'.format(Labelname,file_name) 
                    with open(path,"wb") as f:
                        f.write(file_data)
                    # open(path,"wb") as f: のpathはstrで./内にそのファイルが無くてもpath名で新規作成される。
                    # https://note.nkmk.me/python-file-io-open-with/ ←参考
            except Exception as e: 
                print("Sort_Attachmentで添付ファイルを振り分ける上で『{}』のエラー!".format(e))
                
Sort_Attachment(MessageDataDict)