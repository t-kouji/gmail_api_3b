import openpyxl
from datetime import date,datetime
from glob import glob
import os
import pandas as pd
import numpy as np
import scipy.stats as stats

"""同フォルダ内のフォルダ名をリストで取得"""
def get_dir_name():
    path = "./projects"
    files = os.listdir(path)
    dir_name = [f for f in files if os.path.isdir(os.path.join(path, f))]
    return dir_name

"""前回シートオブジェクトを取得"""
def get_before_ws():
    date_ = date(2000,1,1) #基準となる日付
    index = 0
    for ind,sheet in enumerate(ws_name_l): #シート名一覧をforループしインデックスを取得
        for cell in wb[sheet]['A']: #各シートのA列のセルを上から順番に検索
            if cell.value == '今回検針日：': #もし'今回検針日：'の文字列があれば・・・
                try:
                    d = cell.offset(row=0,column=1).value.date() #右となりのセルを日付データとして取得
                except AttributeError as e:
                    print("{}ファイルの今回検針日の取得に際し『{}』のエラー".format(filename,e))
                if d > date_:
                    date_ = d #dが最新ならdate_を更新していく
                    index = ind #それと同時にindexも更新していく
                break #'今回検針日'の文字列発見し、上記日付更新できたらforループの2ネスト目をbreak
    before_sheet = wb.worksheets[index] #直近のワークシートをbefore_sheetとして取得
    before_date = date_ #日付をbefore_dateとして取得
    return before_sheet,before_date

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
def get_df(dir_name):
    csv_list = glob("./projects/{}/data/*csv".format(dir_name)) #CSVが入っているフォルダ内dataフォルダ内のCSVファイル名をリストで取得
    counter = 0
    csv_name = 0
    for n in csv_list:
        try:
            name = os.path.basename(n) #csvファイル名取得
            c = int(name.rsplit('(')[1].split('_')[0]) #csvファイル名の日付部分をintとして抽出
            if c > counter: #日付が最新のものを選定
                counter = c
                csv_name = name
        except Exception as e:
            print("{}フォルダで最新日付のcsvファイルを抽出する上で『{}』のエラー".format(dir_name,e))
    df = pd.read_csv("./projects/{}/data/{}".format(dir_name,csv_name),encoding='cp932',
    usecols=lambda x : x not in ['Date','Time'],skiprows=[1,2,3,4,5,6],na_values='-') #日時を除外してデータフレームを作成
    median_df = df.median() #dfの中央値を取得。
    df = df.fillna(median_df) #欠損値を中央値に置き換える。
    latest_date = datetime.strptime(str(counter),"%y%m%d").date() #CSVの最新日付を取得（ファイル名より）。datetime → date型へ
    return df,latest_date

"""スミルノフ－グラブス検定を用いdfから外れ値を除去した後に列毎の最大値をリストで取得"""
def get_max_list(df_, alpha=0.01):
    max_list = [] 
    try:
        for columun_name,item in df_.iteritems(): #iteritems()メソッドを使うと、1列ずつコラム名（列名）とその列のデータ（pandas.Series型）を取得できる。
            x, o = list(item), []
            while True:
                n = len(x)
                t = stats.t.isf(q=(alpha / n) / 2, df=n - 2)
                tau = (n - 1) * t / np.sqrt(n * (n - 2) + n * t * t)
                i_min, i_max = np.argmin(x), np.argmax(x)
                myu, std = np.mean(x), np.std(x, ddof=1)
                i_far = i_max if np.abs(x[i_max] - myu) > np.abs(x[i_min] - myu) else i_min
                tau_far = np.abs((x[i_far] - myu) / std)
                if tau_far < tau: break
                o.append(x.pop(i_far))
            max_list.append(np.array(x).max()) #外れ値除外したseriesの中の最大値をmax_listへ追加
            # if not np.array(o).size == 0:
            #     print("外れ値と判断し除外した数字は ",np.array(o))
        return (max_list)
    except Exception as e:
        print("{}maxリストを作成する際に『{}』のエラー".format(filename,e))

"""新シートへの書き込み"""
def writting():
    new_sheet['B4'].value = "={}!B5".format(before_sheet.title) #前回検針日の変更
    new_sheet['B5'].value = latest_date #今回検針日を変更
    for cells in new_sheet.rows: #新しいシートの行で検索。rowsはopenxlのメソッドでtuple型を返す。
        for cell in cells:
            if cell.value ==("積算\n運転時間計\n前回値" or "積算\n運転時間計\n初回値") :
                for c in range(count):                    
                    cell.offset(row=c+3,column=0).value = before_list[0][c] #上記文字列から下にc番目のセルへ値を書き込む
                # break
            if cell.value ==("積算\n電力量計\n前回値" or "積算\n電力量計\n初回値") :
                for c in range(count):                    
                    cell.offset(row=c+3,column=0).value = before_list[1][c] #上記文字列から下にc番目のセルへ値を書き込む
                # break
    # df_list = df_tail.values.tolist() #データフレームをリストへ変換
    cc = new_sheet[criteria_cell]
    counter = 0
    for i in max_list:
        try:
            cc.offset(row=0,column=counter).value = float(i)
            counter += 1
        except Exception as e:
            print("{}ファイルの新シートに書き込みする上で『{}』のエラー".format(filename,e))

dir_list = get_dir_name() #projectsフォルダ内にあるフォルダをリストへ
for dir_name in dir_list:
    filename_list = glob('./projects/{}/*.xlsx'.format(dir_name)) #projectsフォルダ内の各フォルダにあるxlsxファイル名を取得しリストへ。
    #注）各フォルダに置く効果検証xlsxファイルは１つでないといけない。
    try :
        filename = filename_list[0] #globで取得できるものはリストなので変数化。xlsxファイルは一個という前提で[0]。
        wb = openpyxl.load_workbook(filename)
    except IndexError:
        print("{}フォルダに効果検証Excelフォーマットがありますか？".format(dir_name))
        continue
    ws_name_l = wb.sheetnames #シート名をリストとして取得
    criteria_cell = "s5" #積算値記載の基準となるセル番地＝機械台数カウントの基準となるセル番地
    df,latest_date = get_df(dir_name) #dfと、近日付データを取得
    max_list = get_max_list(df) #スミルノフ－グラブス検定を用いdfから外れ値を除去した後に列毎の最大値をリストで取得
    print(max_list)
    before_sheet,before_date = get_before_ws() #前回シート取得と、前回日付データを取得
    if latest_date == before_date:
        print("{}の前回の検証日と今回のcsvの日付が同じではありませんか？".format(filename))
        continue
    count = get_count() #機械台数カウント
    new_sheet = create_new_ws() #新シートオブジェクトを作成
    before_list = get_before_number() #前回シートの必要値を取得しリストへ
    writting() #新シートへの書き込み
    wb.save(filename) #保存