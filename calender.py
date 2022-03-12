import pandas as pd
import numpy as np
from pandas.core.frame import DataFrame
import streamlit as st
import openpyxl
from streamlit.state.session_state import Value
import datetime
import urllib
import urllib.request
import glob
import tabula
import base64
from io import BytesIO

st.set_page_config(page_title='納期カレンダー作成')
st.markdown('#### 納期カレンダー作成')

public_holiday_csv_url="https://www8.cao.go.jp/chosei/shukujitsu/syukujitsu.csv"
public_holiday2 = public_holiday_csv_url.split("/")[-1] #ファイル名の取り出し

# 対象年度
target_year = 2022

# 会社特有の休日
company_holiday = [2022/4/29, 2022/4/30, 2022/5/2, 2022/8/12, 2022/8/13, 2022/8/15, 2022/8/16]

kadoubi = []
chakubi = []

def get_kadoubi(date):
    # 0埋め解消　祝日ファイルに合わせて
    year = date.strftime("%Y")
    month = date.strftime("%m").lstrip("0") #strの左から0を削除
    day = date.strftime("%d").lstrip("0")

    # 日曜日
    if (date.weekday() == 6): # 0月曜 6日曜
        return False
    
    # 祝日
    holidays_df = pd.read_table(public_holiday2, delimiter=',', encoding="SHIFT-JIS")
    if date.strftime(year + "/" + month + "/" + day) in holidays_df['国民の祝日・休日月日'].tolist():
        return False 

    #　会社の休日
    if date.strftime(year + "/" + month + "/" + day) in company_holiday:
        return False

    return True 

def get_chakubi(date):
    # 0埋め解消　祝日ファイルに合わせて
    year = date.strftime("%Y")
    month = date.strftime("%m").lstrip("0")
    day = date.strftime("%d").lstrip("0")

    # 日曜日
    if (date.weekday() == 6): # 0月曜 6日曜
        return False
    
    # 水曜
    if (date.weekday() == 2): # 0月曜 6日曜
        return False    

    # 祝日
    holidays_df = pd.read_table(public_holiday2, delimiter=',', encoding="SHIFT-JIS")
    if date.strftime(year + "/" + month + "/" + day) in holidays_df['国民の祝日・休日月日'].tolist():
        return False

    #　会社の休日
    if date.strftime(year + "/" + month + "/" + day) in company_holiday:
        return False

    return True

    
def generate_pdf():
    # ***ファイルアップロード 今期***
    uploaded_file = st.file_uploader('出荷日表PDFの読み込み', type='pdf', key='shukka')
    df = DataFrame()
    if not uploaded_file:
        st.info('出荷日表PDFを選択してください。')
        st.stop() 
    elif uploaded_file:
        df = tabula.read_pdf(uploaded_file, lattice=True) #dfのリストで出力される
       

    #表が格子状になっている場合 lattice=True そうでない　stream=True　複数ページ読み込み pages='all'
    df_calend = df[0]
    df_calend = df_calend.dropna(how='any')
    df_calend = df_calend.drop(df_calend.columns[[5, 6, 7]], axis=1) #40日から右カラムの削除
    df_calend = df_calend.rename(columns={'Unnamed: 0': '受注日', 'KX250AX\rKX260AX': 'SEOTO-EX'})

    # df[0] = df[0].dropna(how='any')
    # df[0] = df[0].drop(df[0].columns[[5, 6, 7]], axis=1) #40日から右カラムの削除
    # df[0] = df[0].rename(columns={'Unnamed: 0': '受注日', 'KX250AX\rKX260AX': 'SEOTO-EX'})

    #曜日を消す
    # df_calend = df[0]
    df_calend['Aパターン'] = df_calend['Aパターン'].str[:-2]
    df_calend['Bパターン'] = df_calend['Bパターン'].str[:-2]
    df_calend['30日'] = df_calend['30日'].str[:-2]

    #2022年の追加
    df_calend['SEOTO-EX'] = '2022年' + df_calend['SEOTO-EX']
    df_calend['Aパターン'] = '2022年' + df_calend['Aパターン']
    df_calend['Bパターン'] = '2022年' + df_calend['Bパターン']
    df_calend['30日'] = '2022年' +df_calend['30日']

    #datetime型に変換
    df_calend['SEOTO-EX'] = pd.to_datetime(df_calend['SEOTO-EX'], format='%Y年%m月%d日')
    df_calend['Aパターン'] = pd.to_datetime(df_calend['Aパターン'], format='%Y年%m月%d日')
    df_calend['Bパターン'] = pd.to_datetime(df_calend['Bパターン'], format='%Y年%m月%d日')
    df_calend['30日'] = pd.to_datetime(df_calend['30日'], format='%Y年%m月%d日')

    #　str型へ　時間を消す
    # Series型にはdtというアクセサが提供されており、Timestamp型を含む日時型の要素から、日付や時刻のみの要素へ一括変換できます。
    df_calend['SEOTO-EX'] = df_calend['SEOTO-EX'].dt.strftime('%Y-%m-%d')
    df_calend['Aパターン'] = df_calend['Aパターン'].dt.strftime('%Y-%m-%d')
    df_calend['Bパターン'] = df_calend['Bパターン'].dt.strftime('%Y-%m-%d')
    df_calend['30日'] = df_calend['30日'].dt.strftime('%Y-%m-%d')
    
    # radio button
    day_list = [1, 2, 3, 4, 5, 6, 7]
    option_day = st.radio(
     "出荷日の次の日から何日後を着日とするか？（稼働日）",
     day_list, index=4
    )

    arrival_ex = []
    arrival_a = []
    arrival_b = []
    arrival_30 = []

    i = 0
 
    for b1 in df_calend['SEOTO-EX']:
        idx = kadoubi.index(b1) #list内の順番を検索抽出
        arrival_ex_culc = kadoubi[idx + option_day] #着日算出
        if arrival_ex_culc in chakubi:
            arrival_ex.append(arrival_ex_culc)
        else:
            while arrival_ex_culc not in chakubi:
                i += 1
                arrival_ex_culc = kadoubi[idx + option_day + i]  
            arrival_ex.append(arrival_ex_culc)
            i = 0


    for b2 in df_calend['Aパターン']:
        idx = kadoubi.index(b2) #list内の順番を検索抽出
        arrival_a_culc = kadoubi[idx + option_day] #着日算出
        if arrival_a_culc in chakubi:
            arrival_a.append(arrival_a_culc)
        else:
            while arrival_a_culc not in chakubi:
                i += 1
                arrival_a_culc = kadoubi[idx + option_day + i]
            arrival_a.append(arrival_a_culc)
            i = 0

    for b3 in df_calend['Bパターン']:
        idx = kadoubi.index(b3) #list内の順番を検索抽出
        arrival_b_culc = kadoubi[idx + option_day] #着日算出
        if arrival_b_culc in chakubi:
            arrival_b.append(arrival_b_culc)
        else:
            while arrival_b_culc not in chakubi:
                i += 1
                arrival_b_culc = kadoubi[idx + option_day + i]
            arrival_b.append(arrival_b_culc)
            i = 0

    for b4 in df_calend['30日']:
        idx = kadoubi.index(b4) #list内の順番を検索抽出
        arrival_30_culc = kadoubi[idx + option_day] #着日算出
        if arrival_30_culc in chakubi:
            arrival_30.append(arrival_30_culc)
        else:
            while arrival_30_culc not in chakubi:
                i += 1
                arrival_30_culc = kadoubi[idx + option_day + i]
            arrival_30.append(arrival_30_culc)
            i = 0

    arrival_a2 = []
    arrival_b2 = []
    arrival_ex2 = []
    arrival_302 = []

    for c1 in arrival_a:
        c1 = datetime.datetime.strptime (c1, '%Y-%m-%d')
        c1 = c1.strftime('%m/%d')
        arrival_a2.append(c1)
    

    for c2 in arrival_b:
        c2 = datetime.datetime.strptime (c2, '%Y-%m-%d')
        c2 = c2.strftime('%m/%d')
        arrival_b2.append(c2)

    for c3 in arrival_ex:
        c3 = datetime.datetime.strptime (c3, '%Y-%m-%d')
        c3 = c3.strftime('%m/%d')
        arrival_ex2.append(c3)

    for c4 in arrival_30:
        c4 = datetime.datetime.strptime (c4, '%Y-%m-%d')
        c4 = c4.strftime('%m/%d')
        arrival_302.append(c4)

    df_output = pd.DataFrame({
        '受注日': df_calend['受注日'],
        'A (レギュラー)': arrival_a2,
        'B (下記参照)': arrival_b2,
        'SEOTO-EX': arrival_ex2,
        'C（納期30日）': arrival_302,

    })

    def to_excel(df):
        # ***ファイルアップロード フォーム***
        # uploaded_file = st.file_uploader('納期カレンダーフォーム(Excel)の読み込み', type='xlsx', key='form')
        # if not uploaded_file:
        #     st.info('納期カレンダー 書式（Excel）を選択してください。')
        #     st.stop()
        # elif uploaded_file:
        #     workbook = openpyxl.load_workbook(uploaded_file)
        #     sheet = workbook.active #アクティブなワークシートを選択
        #     row_index = 8 #書き込み開始行　1から始まる

        #     for index, rows in df.iterrows():
        #         col_index = 1 #書き込み開始列 1から始まる
        #         for data in rows:
        #             sheet.cell(row=row_index, column=col_index, value=data) #指定したセルにデータの書き込み
        #             col_index = col_index + 1 #書き込む列数をずらす
        #         row_index = row_index + 1 #書き込む行数をずらす
        #     workbook.save("calender1.xlsx")
        # st.download_button(label='Download Excel file', data=sheet.values, file_name= 'calender.xlsx')           


        output = BytesIO()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')
        df_output.to_excel(writer, index = False, sheet_name='Sheet1')
        workbook  = writer.book
        worksheet = writer.sheets['Sheet1']
        format1 = workbook.add_format({'num_format': '0.00'}) # Tried with '0%' and '#,##0.00' also.
        worksheet.set_column('A:A', None, format1) # Say Data are in column A
        writer.save()
        processed_data = output.getvalue()
        return processed_data

  
    # to_excel(df_output)
    df_xlsx = to_excel(df_output)
    st.download_button(label='Download Excel file', data=df_xlsx, file_name= 'calender.xlsx')


    st.markdown('###### GW、お盆、年末年始等が絡む期間は使用を避けてください。')
    # st.markdown(get_table_download_link(df_output), unsafe_allow_html=True)
    st.caption('selected {}'.format(option_day))
    st.caption('5日まで検証済')

if __name__ == '__main__':

    # 内閣府から祝日データを取得、更新したいときにTrueにする
    if True:
        urllib.request.urlretrieve(public_holiday_csv_url, public_holiday2)

    # 稼働日　年始から日付を回す
    date = datetime.datetime(target_year, 1, 1)
    while date.year == target_year:
        if get_kadoubi(date):
            kadoubi.append(date.strftime("%Y-%m-%d"))
        date += datetime.timedelta(days=1)

    # 着日　年始から日付を回す
    date = datetime.datetime(target_year, 1, 1)
    while date.year == target_year:
        if get_chakubi(date):
            chakubi.append(date.strftime("%Y-%m-%d"))
        date += datetime.timedelta(days=1)

    generate_pdf()
