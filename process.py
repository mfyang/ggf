import pandas as pd
import re
from openpyxl import load_workbook

TG_F = '处理团购.xlsx'
WZ_F = '处理网站.xlsx'
ZD_F = '处理总单.xlsx'


print('读入商品名字对应表')
df = pd.read_excel('商品名字对应表.xlsx')
names = dict(zip(df.raw.values, df.name.values))


print('转化团购订单, 并保存为 %s' % TG_F)
tg = pd.read_excel('团购.xlsx',sheet_name=2, header=1)
tg2 = tg.drop(tg[tg['Order id']=='Order id'].index)[:-1]
orders = []
for i,row in tg2.iterrows():
    o = {}
    o['客户'] = row['Name']
    o['电话'] = row['Phone']
    o['地址'] = row['Address']
    if row['Note']:
        o['备注'] = row['Note']
    else:
        o['备注'] = '-'
    o['source'] = '团购'
    for c in tg2.columns[8:]:
        o[c] = row[c]
    orders.append(o)

orders1 = orders
df = pd.DataFrame(orders).fillna('')
df.rename(mapper=names, inplace=True)
df.to_excel(TG_F)


print('转化网站订单, 并保存为 %s' % WZ_F)
ggf = pd.read_excel('ggf.xlsx')
orders = []
for i,row in ggf.iterrows():
    o = {}
    o['客户'] = row['收货人']
    o['电话'] = row['收货人电话']
    o['地址'] = "%s,%s,%s" %(row['收货人地址'], row['收货人所在城市'], row['收货人所在省份'])
    o['支付'] = "%s %s" %(row['支付方式'].strip(), row['支付状态'].strip())
    o['备注'] = row['买家留言']
    o['source'] = '网站'
    for item in re.findall('商品名称:(?P<商品>.*?)规格.*?商品价格:(?P<价格>.*?)购买数量:(?P<数量>\d+)', row['商品信息']):
        o[item[0].strip()] = int(item[-1])
    orders.append(o)

df = pd.DataFrame(orders).fillna('')
df.rename(mapper=names, inplace=True)
df.to_excel(WZ_F)

a = input('把两边订单的商品名称协调一致且合并? 需要把两边订单商品名称协调一致 ')
if a:
    df1 = pd.read_excel(TG_F)
    df2 = pd.read_excel(WZ_F)
    total_df = pd.concat([df1, df2])
    total_df.to_excel(ZD_F)
    print('合并两边的订单， 并保存为 %s' % ZD_F)


print('请将订单分线路，每条线路保存在一个excel文件的一个sheet里')
a = input('分好了? ')
if a:
    with pd.ExcelWriter('线路分单.xlsx') as writer:
        sheets = load_workbook('线路.xlsx')
        for sheet in sheets:
            df = pd.read_excel('线路.xlsx', sheet=sheet)
            phones = df['电话'].values
            total_df[total_df['电话'].isin(phones)].to_excel(writer, sheet_name='%s.xlsx' % sheet)
