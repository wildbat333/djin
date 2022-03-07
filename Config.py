import pandas as pd
from pandas import options
#import xlrd

options.io.excel.xlsx.writer = "xlwt"
file_name = "smart_ivanovo_t.xls"

# table_columns = {'Dates':["01.01.22"], 'Sale_manager':["Долг"],
#                  'Code_shop':["Ч4688Ч4688"], 'Name_shop':["Магазин (Пел"],
#                  'Code_position':["0D1-0700D1-070"], 'Name_position':['Семечки "Д'],
#                  'Count_positions':[9], 'Sum_positions':[111.11]}

tp = ["Долгачева Ирина (ЭК СМАРТ",
      "Дроздова Марина (ЭК СМАРТ",
      "Корсакова Елена (ЭК СМАРТ",
      "Федотова Анна (ЭК СМАРТ)"]
d = ['0D1-035', 'Семечки "Джинн"35г.(100)', 100,
     '0D1-070', 'Семечки "Джинн"70г.(50)', 50,
     '0D1-140', 'Семечки "Джинн"(10%)140г.(25)', 25,
     '0D1-250', 'Семечки "Джинн"250г.(14)', 14,
     '0D1-350', 'Семечки "Джинн"350г.(10)', 10]
d_s = ['0D2-035', 'Семечки "Джинн"(соленые)35г.(100)', 100,
       '0D2-070', 'Семечки "Джинн"(соленые)70г.(50)', 50,
       '0D2-140', 'Семечки "Джинн"(соленые)(10%)140г.(25)', 25,
       '0D2-250', 'Семечки "Джинн"(соленые)250г.(14)', 14,
       '0D2-350', 'Семечки "Джинн"(соленые)350г.(10)', 10]
v = ['0V1-050', 'Семечки "Солнечный Великан"Джинн"50г.(60)', 60,
     '0V1-100', 'Семечки "Солнечный Великан"Джинн"100г.(30)', 30,
     '0V1-120', 'Семечки "Солнечный Великан"Джинн"120г.(25)', 25,
     '0V1-200', 'Семечки "Солнечный Великан"Джинн"200г.(15)', 15]
v_s = ['0V2-050', 'Семечки "Солнечный Великан"Джинн"(соленые)50г.(60)', 60,
       '0V2-100', 'Семечки "Солнечный Великан"Джинн"(соленые)100г.(30', 30,
       '0V2-120', 'Семечки "Солнечный Великан"Джинн"(соленые)120г.(25', 25,
       '0V2-200', 'Семечки "Солнечный Великан"Джинн"(соленые)200г.(15', 15]
v_o_s = ['0V3-100', 'Семечки "Солнечный Великан"Джинн"(особо сол100г(30', 30]
m = ['0M1-070', 'Семечки "Мастер Жарки"70г.(50)', 50,
     '0M1-140', 'Семечки "Мастер Жарки"140г.(25)', 25,
     '0M1-250', 'Семечки "Мастер Жарки"250г.(14)', 14,
     '0M1-350', 'Семечки "Мастер Жарки"350г.(10)', 10]
k = ['0K2-050', 'Арахис "Караван орехов"50г.(40)', 40,
     '0K2-090', 'Арахис "Караван орехов"90г.(22)', 22,
     '0K2-150', 'Арахис "Караван орехов"150г.(20)', 13.3,
     '0K2-500', 'Арахис "Караван орехов"500г.(50)', 4]
k_s = ['0S2-040', 'Арахис "Караван орехов"Стандарт"40г.(50)', 50,
       '0S2-090', 'Арахис "Караван орехов"Стандарт"90г.(22)', 20]
t = ['0T2-050', 'Семечки тыквы "Джинн"(соленые)50г.(70)', 70,
     '0T2-100', 'Семечки тыквы "Джинн"(соленые)100г.(35)', 35]

djin = dict(Name='Семечки "Джинн"70г.(50)', Key='0D1-070', Koef='50')

djin_sol = []

#wb = xlrd.open_workbook(file_name, encoding_override='CORRECT_ENCODING')
df = pd.read_excel(file_name, header=None)
print(df.head(5))

df.info()

summm = df.loc[1, 6]
print(summm)

group_koef = []
group_koef.extend(d)
group_koef.extend(d_s)
group_koef.extend(v)
group_koef.extend(v_s)
group_koef.extend(v_o_s)
group_koef.extend(m)
group_koef.extend(k)
group_koef.extend(k_s)
group_koef.extend(t)
#print(group_koef)

for i, row in df.iterrows():
    for f in range(len(group_koef)):
        if df.loc[i, 4] == group_koef[f]:
            df.loc[i, 8] = group_koef[f+2]
            df.loc[i, 9] = round(df.loc[i, 6] / df.loc[i, 8], 2)
            if df.loc[i, 4] in d: df.loc[i, 10] = 'Джин'
            if df.loc[i, 4] in d_s: df.loc[i, 10] = 'Джин соль'
            if df.loc[i, 4] in v: df.loc[i, 10] = 'Великан'
            if df.loc[i, 4] in v_s: df.loc[i, 10] = 'Великан соль'
            if df.loc[i, 4] in v_o_s: df.loc[i, 10] = 'Великан особо соль'
            if df.loc[i, 4] in m: df.loc[i, 10] = 'Мастер'
            if df.loc[i, 4] in k: df.loc[i, 10] = 'Караван'
            if df.loc[i, 4] in k_s: df.loc[i, 10] = 'Караван стандарт'
            if df.loc[i, 4] in t: df.loc[i, 10] = 'Тыква'
        else:
            f = f+3
print(df.head(15))

agg_func_sum = {9:['sum']}
svod_df = df.groupby([1, 10]).agg(agg_func_sum)

agg_func_sum = {6:['sum']}
prom_df = df.groupby([1, 4, 5]).agg(agg_func_sum)
prom_df.info()
#for i in range(d):

#df = df.loc[~df['STP'].isin([1005092])]
#prom_df1 = prom_df.iloc[:, 1].isin(["Долгачева Ирина (ЭК СМАРТ"])
#print(prom_df1)
d_result = [100, 200, 500, 750]
slov = {'d':d_result}
df_result = pd.DataFrame(slov, index=tp)
print(df_result)
svod_df.to_excel("smart_ivanovo_svod.xls", sheet_name='Сводка', startrow=0, merge_cells=False)
prom_df.to_excel("smart_ivanovo_packs.xls", sheet_name='Пачки', startrow=0, merge_cells=False)
df.to_excel("smart_ivanovo_all.xls", sheet_name='Полная', startrow=0, merge_cells=False)

#df[df[1] == tp[0]]

#print(df.last_valid_index)
#keys = []
#for index in df.itertuples():
#keys[df.columns[4]] = df.columns[5]
    #df.columns[5]
#print(keys)
#df.iterrows()
#n = 0
# for row in df.itertuples(index=False):
#     keys.append(df.columns[4])
#     n = n +1
#     #print(row)
# print(n)
# print(keys)
#for row in df.iterrows():
	#print(f"{row}")
# tps = []
# for row in df.itertuples(index=False):
#     tps.append(df.columns[1])
# print(tps)
# tp_b = list(set(tps))
# print(tp_b)
