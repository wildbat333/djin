import pandas as pd
from openpyxl import Workbook

file_name_input = "smart_ivanovo.xlsx"
file_name_result = "smart_ivanovo_result.xlsx"

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
       '0V2-200', 'Семечки "Солнечный Великан"Джинн"(соленые)200г.(15', 15,
       '0V2-300', 'Семечки "Солнечный Великан"Джинн"(соленые)300г.(10', 10]
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
       '0S2-090', 'Арахис "Караван орехов"Стандарт"90г.(22)', 22]
t = ['0T2-050', 'Семечки тыквы "Джинн"(соленые)50г.(70)', 70,
     '0T2-100', 'Семечки тыквы "Джинн"(соленые)100г.(35)', 35]

wb = Workbook()
ws = wb.active
ws1 = wb.create_sheet('Сводка')
ws2 = wb.create_sheet('Пачки')
ws3 = wb.create_sheet('Полная')
wb.remove(ws)
wb.save(file_name_result)


df = pd.read_excel(file_name_input, header=None)

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


for i, row in df.iterrows():
    for f in range(len(group_koef)):
        if df.loc[i, 4] == group_koef[f]: #Можно искать по названию в [i, 5], тогда f=1 в начале
            df.loc[i, 8] = group_koef[f+2] # Тогда тут f+1
            df.loc[i, 9] = df.loc[i, 6] / df.loc[i, 8]
            if df.loc[i, 4] in d: df.loc[i, 10] = 'Джин' #тут везде ставим поиск тогда в [i, 5]
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


agg_func_sum = {6:['sum']}
prom_df = df.groupby([1, 4, 5]).agg(agg_func_sum) #Тут тогда 4 надо исключать, будет только по 5

svod_df = pd.crosstab(df[1],
                      df[10],
                      values=df[9],
                      aggfunc='sum',
                      normalize=False)
svod_df = svod_df.round(2)


with pd.ExcelWriter(file_name_result, mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
    svod_df.to_excel(writer, sheet_name='Сводка', startrow=0)
    prom_df.to_excel(writer, sheet_name='Пачки')
    df.to_excel(writer, sheet_name='Полная')



