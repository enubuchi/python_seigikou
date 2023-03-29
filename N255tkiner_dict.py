import os
import sys
import datetime as dt
import pandas_datareader.data as web
import matplotlib.pyplot as plt
import datetime
import tkinter
from tkinter import *
from tkinter import ttk

root = Tk()
root.title('Stock Chart')

# ウィジェットの作成
frame1 = ttk.Frame(root, padding=16)
label1 = ttk.Label(frame1, text='日付(年/月/日で入力)')
t = StringVar()
entry1 = ttk.Entry(frame1, textvariable=t)

# 各種ウィジェットの作成
varhigh = tkinter.IntVar()
varhigh.set(1) 
varlow = tkinter.IntVar()
varlow.set(1) 
varstart = tkinter.IntVar()
varstart.set(1) 
varend = tkinter.IntVar()
varend.set(1) 
highon = ttk.Checkbutton(frame1, text="高値", variable=varhigh)
lowon = ttk.Checkbutton(frame1, text="低値", variable=varlow)
starton = ttk.Checkbutton(frame1, text="始値", variable=varstart)
endon = ttk.Checkbutton(frame1, text="終値", variable=varend)

button1 = ttk.Button(
    frame1,
    text='OK',
    command=root.destroy
    )

# レイアウト
frame1.pack()
highon.pack(side=LEFT)
lowon.pack(side=LEFT)
starton.pack(side=LEFT)
endon.pack(side=LEFT)
label1.pack(side=LEFT)
entry1.pack(side=LEFT)
button1.pack(side=LEFT)

# ウィンドウの表示開始
root.mainloop()
checksum = varhigh.get() + varlow.get()*2 + varstart.get()*4 + varend.get()*8
#print (checksum)

#銘柄コード入力(7177はGMO-APです。)
# 7261マツダ
# 日経平均 998407.O
ticker_symbol="^NKX"
ticker_symbol_dr=ticker_symbol #+ “.JP”
#2022-01-01以降の株価取得
start = datetime.datetime.strptime(t.get(), '%Y/%m/%d')
# start='2022-01-01'
end = dt.date.today()
#データ取得
df = web.DataReader(ticker_symbol_dr, data_source='stooq', start=start,end=end)
#2列目に銘柄コード追加
df.insert(0, "code", ticker_symbol, allow_duplicates=False)
#csv保存
fn = os.path.dirname(__file__)
df.to_csv( os.path.dirname(__file__) + '/s_stock_data_'+ ticker_symbol + '.csv')
#df_stock = df.read()[‘Close’]
#print(df_stock)
#matplotlibで株価をグラフ化

showlist ={1:"High", 2:"Low", 3:["High","Low"], 4:"Open", 5:["High","Open"], 6:["Low","Open"],
           7:["High","Low","Open"], 8:"Close", 9:["High","Close"], 10:["Low","Close"],
           11:["High","Low","Close"], 12:["Open","Close"], 13:["High","Open","Close"],
          14:["Low","Open","Close"], 15:["High","Low","Open","Close"]}

if checksum in showlist:
    show = showlist[checksum]
    df[show].head(90).plot(figsize=(16,8),fontsize=18)

else:
    root2 = Tk()
    root2.title('やり直し')
    frame2 = ttk.Frame(root2, padding=16)
    label2 = ttk.Label(frame2, text='いずれかにチェックを入れて、やり直してください。')
    button2 = ttk.Button(
        frame2,
        text='OK',
        command=root2.destroy
        )
    frame2.pack()
    label2.pack(side=TOP)
    button2.pack(side=BOTTOM)
    root2.mainloop()
    sys.exit()
plt.legend(bbox_to_anchor=(0, 1), loc='upper left', borderaxespad=1, fontsize=18)
plt.grid(True)
plt.title('N225 Graph',fontsize=18)
#plt.ylim(0, 30000)
plt.savefig("N225 graph.png")
plt.show()
