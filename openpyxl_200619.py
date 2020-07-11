import openpyxl
from datetime import date,datetime
from glob import glob
import os
import pandas as pd

"""同フォルダ内のフォルダ名をリストで取得"""
def get_dir_name():
    path = "./projects"
    files = os.listdir(path)
    dir_name = [f for f in files if os.path.isdir(os.path.join(path, f))]
    # print(dir_name)
    return dir_name

"""前回シートオブジェクトを取得"""
def get_before_ws():
    d = date(2000,1,1) #基準となる日付
    index = 0
    for ind,sheet in enumerate(ws_name_l):
        for cell in wb[sheet]['A']: #各シートのA列のセルを上から順番に検索
            if cell.value == '今回検針日：': #もし'今回検針日：'の文字列があれば・・・
                try:
                    date_ = cell.offset(row=0,column=1).value.date() #右となりのセルを日付データとして取得
                except AttributeError:
                    pass
                if date_ > d:
                    d = date_ #date_が最新ならdを更新していく
                    index = ind #それと同時にindexも更新していく
    before_sheet = wb.worksheets[index] #直近のワークシートをbefore_sheetとしてしゅとく
    # print(d,before_sheet)
    return before_sheet

"""機械台数の取得"""
def get_count():
    counter = 0
    cc = before_sheet[criteria_cell] 
    while True:
        if cc.offset(row=0,column=counter).value != None:
            counter += 1
        else:
            break
    count = int(counter / 2)
    return count

"""新シートの作成"""
def create_new_ws():
    wb.copy_worksheet(before_sheet) #前回シートのコピー
    d_str = latest_date.strftime("%y.%m.%d") #シート名とする日付の文字列化
    wb.worksheets[-1].title = d_str #シート名を日付に変更
    new_sheet = wb.worksheets[-1]
    return new_sheet #新シートのオブジェクトとして返す

"""前回シートの必要値を取得"""
def get_before_number():
    cc = before_sheet[criteria_cell]
    before_kwh = [cc.offset(row=0,column=c*2).value for c in range(count)]
    before_h = [cc.offset(row=0,column=c*2+1).value for c in range(count)]
    # print("前回シートの値",[before_h,before_kwh])
    return [before_h,before_kwh]
  

"""CSVから値と直近日付を取得"""
def get_csv_df(d):
    file_list = glob("./projects/{}/data/*csv".format(d)) #CSVが入っているフォルダ内のCSVファイル名をリストで取得
    counter = 0
    file_name = 0
    for n in file_list:
        try:
            name = os.path.basename(n) #ファイル名取得
            c = int(name.rsplit('(')[1].split('_')[0]) #ファイル名の日付部分をintとして抽出
            if c > counter: #日付が最新のものを選定
                counter = c
                file_name = name
        except Exception as e:
            print("ファイルの日付を抽出する上で『{}』のエラー".format(e))
    df = pd.read_csv("./projects/{}/data/{}".format(d,file_name),encoding='cp932',
    usecols=lambda x : x not in ['Date','Time']) #日時を除外してデータフレームを作成
    df_tail = df.tail(1) #データフレームの末行を取得
    latest_date = datetime.strptime(str(counter),"%y%m%d")
    return df_tail,latest_date

"""新シートへの書き込み"""
def writting():
    new_sheet['B4'].value = "={}!B5".format(before_sheet.title) #前回検針日の変更
    new_sheet['B5'].value = latest_date
    for cells in new_sheet.rows: #新しいシートの行で検索
        for cell in cells:
            if cell.value ==("積算\n運転時間計\n前回値" or "積算\n運転時間計\n初回値") :
                for c in range(count):                    
                    cell.offset(row=c+3,column=0).value = before_list[0][c] #上記文字列から下にc番目のセルへ値を書き込む
                # break
            if cell.value ==("積算\n電力量計\n前回値" or "積算\n電力量計\n初回値") :
                for c in range(count):                    
                    cell.offset(row=c+3,column=0).value = before_list[1][c] #上記文字列から下にc番目のセルへ値を書き込む
                # break
    df_list = csv_df.values.tolist() #データフレームをリストへ変換
    # print(df_list)
    cc = new_sheet[criteria_cell]
    counter = 0
    for i in df_list[0]:
        try:
            cc.offset(row=0,column=counter).value = float(i)
            counter += 1
        except Exception as e:
            print("ファイルの日付を抽出する上で『{}』のエラー".format(e))

dir_list = get_dir_name() #projectsフォルダ内にあるフォルダをリストへ
for d in dir_list:
    filename_list = glob('./projects/{}/*.xlsx'.format(d)) #projectsフォルダ内の各フォルダにあるxlsxファイル名を取得しリストへ
    try :
        wb = openpyxl.load_workbook(filename_list[0])
    except IndexError:
        print("{}フォルダに効果検証Excelフォーマットがありますか？".format(d))
        continue
    ws_name_l = wb.sheetnames #シート名をリストとして取得
    criteria_cell = "s5" #積算値記載の基準となるセル番地＝機械台数カウントの基準となるセル番地
    csv_df,latest_date = get_csv_df(d) #CSVの各値をdfとして取得し、直近日付データも取得
    before_sheet = get_before_ws() #前回シート取得
    count = get_count() #機械台数カウント
    new_sheet = create_new_ws() #新シートオブジェクトを作成
    before_list = get_before_number() #前回シートの必要値を取得しリストへ
    writting() #新シートへの書き込み
    wb.save(filename_list[0]) #保存