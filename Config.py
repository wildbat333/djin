import pandas as pd
from openpyxl import Workbook
import Constants


wb = Workbook()
ws = wb.active
ws1 = wb.create_sheet('Сводка 2.0')
ws2 = wb.create_sheet('Корректировка')
ws3 = wb.create_sheet('Сводка')
ws4 = wb.create_sheet('Пачки')
ws5 = wb.create_sheet('Полная')
ws6 = wb.create_sheet('СВ')
ws7 = wb.create_sheet('Денюшки')
wb.remove(ws)
wb.save(Constants.file_name_result)


df = pd.read_excel(Constants.file_name_input, header=None)
df_2 = pd.read_excel(Constants.file_name_input_2, header=None)

cel_k = []
cel_e = []

itog_k = []
itog_e = []


clmns = list(df_2)
#print(clmns)
for c in range(len(clmns)):     #Поиск заголовка Итого в 12 строке - шапка
    if df_2.loc[Constants.row_with_itogo, c] == 'Итого': cL_itogo = c
for g, row in df_2.iterrows():
   # if df_2.loc[g, 0] == 'Долгачева Ирина': cel_k.append(g)
    if df_2.loc[g, 0] == Constants.tp[0]: cel_k.append(g)
    if df_2.loc[g, 0] == Constants.tp[1]: cel_k.append(g)
    if df_2.loc[g, 0] == Constants.tp[2]: cel_k.append(g)
    if df_2.loc[g, 0] == Constants.tp[3]: cel_k.append(g)

#Проверка не соскочил ли Итого
if df_2.loc[Constants.row_with_itogo, cL_itogo] == 'Итого':
    print('Количество строк не изменилось, Итого на месте!')
else:
    print('Количество строк изменилось! Сейчас в ячейке Итого: ', df_2.loc[Constants.row_with_itogo, cL_itogo],
          '\nИзмени номер строки!')

#Можно проверить какие ТП забираются из отчета СВ
# tp_from_sv_k = []
# for i in range(len(cel_k)):
#     tp_from_sv_k.append(df_2.loc[cel_k[i], 0])
#print(tp_from_sv_k)

# Отбрасываем строки Групп из таблицы по количествам товаров из констант сверху
for n in range(len(cel_k)):
    for i in range(Constants.number_of_products + 2):
        if i == Constants.number_of_products_in_group_1:
            continue
        elif i == Constants.number_of_products_in_group_1 + Constants.number_of_products_in_group_2 + 1:
            continue
        elif i == Constants.number_of_products + 2:
            break
        else:
            itog_k.append(cel_k[n]+(i+1))


#Можно проверить какие позиции забираются из отчета СВ
products_groups_sv_k = []
for i in range(len(itog_k)):
    products_groups_sv_k.append(df_2.loc[itog_k[i], 0])
# print(products_groups_sv_k)

# Формируем список товаров для заголовков (полный)
columns = []
columns.extend(products_groups_sv_k[:Constants.number_of_products])

tp0 = []
tp0.append(Constants.tp[0])
tp1 = []
tp1.append(Constants.tp[1])
tp2 = []
tp2.append(Constants.tp[2])
tp3 = []
tp3.append(Constants.tp[3])
for i in range(len(itog_k)):
    if i < Constants.number_of_products:
        tp0.append(df_2.loc[itog_k[i], cL_itogo])
    elif i < Constants.number_of_products * 2:
        tp1.append(df_2.loc[itog_k[i], cL_itogo])
    elif i < Constants.number_of_products * 3:
        tp2.append(df_2.loc[itog_k[i], cL_itogo])
    elif i < Constants.number_of_products * 4:                #Это если 4 ТП иначе уменьшать или увеличивать
        tp3.append(df_2.loc[itog_k[i], cL_itogo])

df_3 = pd.read_excel(Constants.file_name_input_3, header=None)


clmns = list(df_3)
#print(clmns)

for c in range(len(clmns)):     #Поиск заголовка Итого в 12 строке - шапка
    if df_3.loc[Constants.row_with_itogo, c] == 'Итого': cL_itogo = c
for g, row in df_3.iterrows():
    if df_3.loc[g, 0] == Constants.tp[4]: cel_e.append(g)
    if df_3.loc[g, 0] == Constants.tp[5]: cel_e.append(g)
    if df_3.loc[g, 0] == Constants.tp[6]: cel_e.append(g)
    if df_3.loc[g, 0] == Constants.tp[7]: cel_e.append(g)

#Проверка не соскочил ли Итого
#print(df_3.loc[Constants.row_with_itogo, cL_itogo])  ##Итого

#Можно проверить какие ТП забираются из отчета СВ
# tp_from_sv_e = []
# for i in range(len(cel_e)):
#     tp_from_sv_e.append(df_2.loc[cel_e[i], 0])
#print(tp_from_sv_e)

# Отбрасываем строки Групп из таблицы по количествам товаров из констант сверху
for n in range(len(cel_e)):
    for i in range(Constants.number_of_products+2):
        if i == Constants.number_of_products_in_group_1:
            continue
        elif i == Constants.number_of_products_in_group_1 + Constants.number_of_products_in_group_2 + 1:
            continue
        elif i == Constants.number_of_products + 2:
            break
        else:
            itog_e.append(cel_e[n]+(i+1))

#Можно проверить какие позиции забираются из отчета СВ
# products_groups_sv_e = []
# for i in range(len(itog_e)):
#     products_groups_sv_e.append(df_3.loc[itog_e[i], 0])
# # print(products_groups_sv_e)

tp4 = []
tp4.append(Constants.tp[4])
tp5 = []
tp5.append(Constants.tp[5])
tp6 = []
tp6.append(Constants.tp[6])
tp7 = []
tp7.append(Constants.tp[7])
for i in range(len(itog_e)):
    if i < Constants.number_of_products:
        tp4.append(df_3.loc[itog_e[i], cL_itogo])
    elif i < Constants.number_of_products * 2:
        tp5.append(df_3.loc[itog_e[i], cL_itogo])
    elif i < Constants.number_of_products * 3:
        tp6.append(df_3.loc[itog_e[i], cL_itogo])
    elif i < Constants.number_of_products * 4:                #Это если 4 ТП иначе уменьшать или увеличивать
        tp7.append(df_3.loc[itog_e[i], cL_itogo])


del tp0[0] #удаляем Имя ТП из списка
del tp1[0]
del tp2[0]
del tp3[0]
del tp4[0]
del tp5[0]
del tp6[0]
del tp7[0]

df_4 = pd.DataFrame([tp0, tp1, tp2, tp3, tp4, tp5, tp6, tp7],
                    columns=columns,
                    index=Constants.tp)
#print(df_4)

group_koef = []
group_koef.extend(Constants.d)
group_koef.extend(Constants.d_s)
group_koef.extend(Constants.v)
group_koef.extend(Constants.v_s)
group_koef.extend(Constants.v_o_s)
group_koef.extend(Constants.m)
group_koef.extend(Constants.k)
group_koef.extend(Constants.k_s)
group_koef.extend(Constants.t)


for i, row in df.iterrows():
    #if df.loc[i, 3] in Dolgacheva: df.loc[i, 1] = tp[0] #убрал 1и2 If, сместил индексы в tp -1
    #if df.loc[i, 1] == tp_1c[0]: df.loc[i, 1] = tp[0]
    if df.loc[i, 3] in Constants.Drozdova: df.loc[i, 1] = Constants.tp[0]
    if df.loc[i, 1] == Constants.tp_1c[0]: df.loc[i, 1] = Constants.tp[0]
    if df.loc[i, 3] in Constants.Korsakova: df.loc[i, 1] = Constants.tp[1]
    if df.loc[i, 1] == Constants.tp_1c[1]: df.loc[i, 1] = Constants.tp[1]
    if df.loc[i, 3] in Constants.Krasnova: df.loc[i, 1] = Constants.tp[4]
    if df.loc[i, 1] == Constants.tp_1c[4]: df.loc[i, 1] = Constants.tp[4]
    if df.loc[i, 3] in Constants.Shutova: df.loc[i, 1] = Constants.tp[7]
    if df.loc[i, 1] == Constants.tp_1c[7]: df.loc[i, 1] = Constants.tp[7]
    if df.loc[i, 3] in Constants.Migushova: df.loc[i, 1] = Constants.tp[5]
    if df.loc[i, 1] == Constants.tp_1c[5]: df.loc[i, 1] = Constants.tp[5]
    if df.loc[i, 3] in Constants.Fedotova: df.loc[i, 1] = Constants.tp[3]
    if df.loc[i, 1] == Constants.tp_1c[3]: df.loc[i, 1] = Constants.tp[3]
    if df.loc[i, 3] in Constants.Chepa: df.loc[i, 1] = Constants.tp[6]
    if df.loc[i, 1] == Constants.tp_1c[6]: df.loc[i, 1] = Constants.tp[6]
    if df.loc[i, 3] in Constants.Ogurova: df.loc[i, 1] = Constants.tp[2]
    if df.loc[i, 1] == Constants.tp_1c[2]: df.loc[i, 1] = Constants.tp[2]
    for f in range(len(group_koef)):
        if df.loc[i, 4] == group_koef[f]: #Можно искать по названию в [i, 5], тогда f=1 в начале
            df.loc[i, 8] = group_koef[f+2] # Тогда тут f+1
            df.loc[i, 9] = df.loc[i, 6] / df.loc[i, 8]
            if df.loc[i, 4] in Constants.d:
                df.loc[i, 10] = '1_Премиум' #тут везде ставим поиск тогда в [i, 5]
                df.loc[i, 11] = '1_Премиум'
            if df.loc[i, 4] in Constants.d_s:
                df.loc[i, 10] = '2_Премиум соль'
                df.loc[i, 11] = '1_Премиум'
            if df.loc[i, 4] in Constants.v:
                df.loc[i, 10] = '3_Полосатая'
                df.loc[i, 11] = '1_Премиум'
            if df.loc[i, 4] in Constants.v_s:
                df.loc[i, 10] = '4_Полосатая соль'
                df.loc[i, 11] = '1_Премиум'
            if df.loc[i, 4] in Constants.v_o_s:
                df.loc[i, 10] = '5_Полосатая особо соль'
                df.loc[i, 11] = '1_Премиум'
            if df.loc[i, 4] in Constants.m:
                df.loc[i, 10] = '6_Мастер'
                df.loc[i, 11] = '2_Тыква'
            if df.loc[i, 4] in Constants.k:
                df.loc[i, 10] = '8_Караван'
                df.loc[i, 11] = '3_Орехи'
            if df.loc[i, 4] in Constants.k_s:
                df.loc[i, 10] = '9_Караван стандарт'
                df.loc[i, 11] = '3_Орехи'
            if df.loc[i, 4] in Constants.t:
                df.loc[i, 10] = '7_Тыква'
                df.loc[i, 11] = '2_Тыква'
        else:
            f = f+3

products_in_svod = ['1_Премиум', '2_Премиум соль', '3_Полосатая', '4_Полосатая соль', '5_Полосатая особо соль',
                    '7_Тыква', '8_Караван', '9_Караван стандарт']


agg_func_sum = {6:['sum']}
prom_df = df.groupby([1, 4, 5]).agg(agg_func_sum) #Тут тогда 4 надо исключать, будет только по 5

svod_df = pd.crosstab(df[1],
                      df[10],
                      values=df[9],
                      aggfunc='sum',
                      normalize=False)
svod_df = svod_df.round(2)
svod_df.index.name = None
svod_df.rename(columns={'1_Премиум': 'Премиум', '2_Премиум соль': 'Премиум соль', '3_Полосатая': 'Полосатая',
                        '4_Полосатая соль': 'Полосатая соль', '5_Полосатая особо соль': 'Полосатая особо соль',
                        '7_Тыква': 'Тыква', '8_Караван': 'Караван', '9_Караван стандарт': 'Караван стандарт'},
               inplace=True)
svod_sort_df = svod_df.loc[['Доставка ОПТ', 'Дроздова Марина', 'Корсакова Елена', 'Огурова Ольга', 'Федотова Анна',
                            'Краснова Наталья', 'Мигушова Надежда', 'Чепа Елена', 'Шутова Ольга', 'Редьков Алексей',
                            'Сотрудники']] #От этой причесанности может сломаться, если не будет какого нибудь индекса

money_df = pd.crosstab(df[1],
                       df[11],
                       values=df[7],
                       aggfunc='sum',
                       normalize=False)
money_df = money_df.round(2)
money_df.index.name = None
money_df.rename(columns={'1_Премиум': 'Премиум', '2_Тыква': 'Тыква', '3_Орехи': 'Орехи'},
                inplace=True)
money_sort_df = money_df.loc[['Доставка ОПТ', 'Дроздова Марина', 'Корсакова Елена', 'Огурова Ольга', 'Федотова Анна',
                              'Краснова Наталья', 'Мигушова Надежда', 'Чепа Елена', 'Шутова Ольга', 'Редьков Алексей',
                              'Сотрудники']]

# print(svod_df.index)
# print(df_4.index)


del tp0[5], tp0[-1] #удаляем Мастер Жарки и Арахис Джинн из списка
del tp1[5], tp1[-1]
del tp2[5], tp2[-1]
del tp3[5], tp3[-1]
del tp4[5], tp4[-1]
del tp5[5], tp5[-1]
del tp6[5], tp6[-1]
del tp7[5], tp7[-1]
#test = tp0
del columns[5], columns[-1]
df_5 = pd.DataFrame([tp0, tp1, tp2, tp3, tp4, tp5, tp6, tp7],
                    columns=columns,
                    index=Constants.tp)

df_6 = svod_df.loc[Constants.tp]
#print(df_6)
df_7 = pd.DataFrame(df_6.values - df_5.values, df_5.index, columns)
# print(df_7)



with pd.ExcelWriter(Constants.file_name_result, mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
    df_6.to_excel(writer, sheet_name='Сводка 2.0', startrow=0)
    df_7.to_excel(writer, sheet_name='Корректировка')
    svod_sort_df.to_excel(writer, sheet_name='Сводка', startrow=0)
    prom_df.to_excel(writer, sheet_name='Пачки')
    df.to_excel(writer, sheet_name='Полная')
    df_4.to_excel(writer, sheet_name='СВ')
    money_sort_df.to_excel(writer, sheet_name='Денюшки')



