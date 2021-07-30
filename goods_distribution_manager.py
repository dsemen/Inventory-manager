from tkinter import *
from tkinter.ttk import *
import pandas as pd
import numpy as np
import random

#root window
window = Tk()
window.title('Помощник менеджера по товародвижению')
window.geometry('810x610')

# notebook
n_book = Notebook()
n_book.pack(padx=10, pady=10, expand=1)

#STOCK FRAME
stock = Frame(n_book, width=700, height=600 )
stock.pack(expand=True, fill=Y)

lbl_stock = Label(stock, text='Отчет по остаткам и товарам в заборных листах')
lbl_stock.pack(pady=10)

v = IntVar()

# Dictionary to create multiple buttons
stock_dict = {"Обручальные кольца": 0, 
              "Буквы и Зодиаки": 1,
              "Товы": 2, 
              "Тип 1 и ТН (для топ100)": 3, 
              "Артикул и Размер":  4 }
 
for (text, value) in stock_dict.items():
    Radiobutton(stock, text = text, variable = v,
        value = value).pack(pady = 7, padx=280, anchor='w')

stock_lst = [
    ["КодСклада", "ТоварНо", "Дизайн", 'ОписаниеТовара', 'ПоставщикАрт','ТолщинаПлетения' , 'ТоварнаяПодгруппа', 'РазмерИзделия'], 
    ["КодСклада", "ТоварНо", 'ТоварноеНаправление', "Дизайн", "Описание3"],
    ["КодСклада", "ТоварНо", "СерийныйНомер"],
    ["КодСклада", "ТоварНо", "СерийныйНомер", "Тип1", "ТоварноеНаправление"],
    ["КодСклада", "ТоварНо", 'ПоставщикАрт', 'РазмерИзделия']
        ]

def stockfunc():
    save_name = list(stock_dict.keys())[v.get()]
    stock_columns = stock_lst[v.get()]
    cols = ["КодСклада", "НазваниеСклада", "ТоварНо","СерийныйНомер", 'ПоставщикАрт', 'Тип1', 'ТолщинаПлетения',
            'Тип3', 'ТоварнаяПодгруппа', "Дизайн", 'ОписаниеТовара',  'ТоварноеНаправление', 'РазмерИзделия', "Проект"]
    if save_name == "Буквы и Зодиаки":
        cols.append("Описание3")
    #загужаем файл остатков и отметаем лишнее
    ost_obr = pd.read_csv(r'C:\анаконда\tmp.csv', delimiter=';', encoding='utf-8',low_memory=False,
                        usecols=cols )
    ost_obr = ost_obr[ost_obr.Проект.isin(['ЗОЛОТОЙ','ОРИОН','ЗОЛОТО','ТАЛАНТ', 'НОУНЕЙМ'])].sort_values('КодСклада')
    
    #делим остатки на две части:магазы и центральные склады (то что на центральных складах может быть уже распределено)
    ost_obr1 = ost_obr[ost_obr.НазваниеСклада.str.contains('центральный склад|опт',case = False)]
    ost_obr = ost_obr[~ost_obr.НазваниеСклада.str.contains('центральный склад|опт',case = False)]
    del ost_obr1['КодСклада']
    
    #обработка заборников верхние пустые строки в загружаемом файле должны быть удалены
    zab = pd.read_excel(r'C:\анаконда\zab.xlsx', skiprows=4)
    zab = zab[["Серийный Но.","Склад-приемник","Номер ЗЛ"]].sort_values(by="Номер ЗЛ",ascending = False)
    zab.columns = ['СерийныйНомер','КодСклада','ЗЛ']
    zab.rename(columns={"Серийный Но.":'СерийныйНомер', "Склад-приемник":'КодСклада',"Номер ЗЛ":'ЗЛ'}, inplace=True) 
    del zab['ЗЛ']
    
    # из заборных листов проставляем номер склада для серийников, которые на остатке 4001/3075/опте
    ost_obr1 = pd.merge(ost_obr1,zab,on=['СерийныйНомер'],how = 'inner')#inner join оставит только неперемещенный товар
    ost_obr1 = ost_obr1[cols]
    
    #соединяем остатки на магазах,на машине, то что числилось на 4001/3075/опте
    result = pd.concat([ost_obr,ost_obr1])
    
    if save_name == "Буквы и Зодиаки":
        result = result[result['Тип1'] == 'ПОДВЕС ДЕКОРАТИВНЫЙ']
    
    result = result[stock_columns]
    result.set_index('КодСклада',drop=True, inplace=True)
      
    #сохраняем в эксель. нужно доработать сохранение в уже существующий эксель
    name = 'C:/Остатки/' + save_name + '.xlsx'
    with pd.ExcelWriter(name) as writer:
        result.to_excel(writer)
    
stock_button = Button(stock, text="Сформировать остатки", command=stockfunc) 
stock_button.pack(pady=27)

# DISTRIBUTION FRAME
distr = Frame(n_book, width=800, height=600)
distr.pack(fill='both', expand=True)

distr_lbl = Label(distr, text='Список магазинов для распределения сохранить в C:\анаконда\shop_list.xlsx (не требуется для букв и зодиаков)')
distr_lbl.pack(pady=7)

distr_lbl = Label(distr, text='накладная для распределения')
distr_lbl.pack()

distr_t2 = Text(distr, width=8, height=1)
distr_t2.pack(pady=7)# поле для файла
#buttons
distr_dict = {"Обручальные кольца": 0, 
              "Буквы и Зодиаки": 1,
              "Любой товар": 2}

var_distr =  IntVar()
for (text, value) in distr_dict.items():
    Radiobutton(distr, text = text, variable = var_distr,
        value = value).pack(side=TOP , ipady=5)

# checkboxes
var_cb1 = BooleanVar()
var_cb1.set(False)
var_cb2 = BooleanVar()
var_cb2.set(False)

    
distr_cb1 = Checkbutton(distr, text='отправить на 7171', variable=var_cb1)
distr_cb2 = Checkbutton(distr, text='убрать нули из накладной', variable=var_cb2)

distr_cb1.pack(ipadx=23, ipady=5)
distr_cb2.pack()



def distr_all(file):
    prih = pd.read_excel(file)
    prih.sort_values(by='Товар', inplace=True)
    shops = pd.read_excel('C:\анаконда\shop_list.xlsx', names=['КодСклада'],dtype='int')
    shop_lst = shops.КодСклада.tolist()
    # словарь тов - количество и список товов
    tov_list = prih.Товар.unique().tolist()
    tov_dict = prih.Товар.value_counts()[prih.Товар.unique()].to_dict()
        
    # остатки
    if prih.iloc[0, 3] == 'КОЛЬЦО ОБРУЧАЛЬНОЕ':
        stock = pd.read_excel(r'C:\Остатки\Обручальные кольца.xlsx', usecols=['КодСклада','ТоварНо'])
    else:
        stock = pd.read_excel(r'C:\Остатки\Товы.xlsx', usecols=['КодСклада','ТоварНо'])
    
    stock = stock[stock['КодСклада'].isin(shop_lst) & stock['ТоварНо'].isin(tov_list)]
    
    # таблица остатки товов по складам    
    pivot = pd.pivot_table(stock,index='КодСклада', columns='ТоварНо', aggfunc=len, fill_value=0).reset_index()
    pivot = shops.merge(pivot, how='left')
    
    # добавим в таблицу колонки с нулями по товам, которых не хватает
    add_cols =  [x for x in tov_list if x not in pivot.columns.to_list()]
    add_zeroes = [list(np.zeros(len(shop_lst), dtype='int'))] * len(add_cols)
    
    pivot = pivot.reindex(columns=pivot.columns.tolist() + add_cols)
    
    pivot.fillna(0, inplace=True)
    pivot = pivot.astype('int')
    pivot = pivot[['КодСклада'] + tov_list]
    
    # Если количество по тову в прихе больше чем в списке то к списку добавить нули, если нет то обрезать
    sklad_list = [
        pivot[pivot[x] == 0]['КодСклада'].tolist() + 
        list(np.zeros((tov_dict[x] - len(pivot[pivot[x] == 0]['КодСклада'].tolist())), dtype='int')) 
        if 
        tov_dict[x] > len(pivot[pivot[x] == 0]['КодСклада'].tolist()) 
        else pivot[pivot[x] == 0]['КодСклада'].tolist()[:tov_dict[x]] 
        for x in tov_list
    ]
    
    sklad_list = [item for sublist in sklad_list for item in sublist]
    
    prih['Период реализации'] = sklad_list
    global distr_cb1, distr_cb2 
    if var_cb1.get() == True:
       prih['Период реализации'] = prih['Период реализации'].apply(lambda x: 7171 if x == 0 else x)
    
    if var_cb2.get() == True:
        prih = prih[prih['Период реализации'] > 0]
    
    name = 'C:/рушники/' + str(random.randrange(5000,10000)) + '.xlsx'
    prih.to_excel(name, sheet_name='Движение товара', index=False)     
        
    
def distr_obr(file):
    shops = pd.read_excel('C:\анаконда\shop_list.xlsx', dtype='int')
    
    # достаем Тпг+размер c количеством которые нужно считать из накладной 
    prih = pd.read_excel(file, dtype='str')
    prih['tpg'] = prih['Товарная подгруппа'].str.cat(prih['Размер'], sep="_")
    prih['tpg'] = prih['tpg'].str.replace('.', ',')# меняем точки на запятые
    prih.sort_values(by=['tpg'], inplace=True)
    tpg = prih['tpg'].unique().tolist()
    tpg.sort()
    tpg_dic = prih['tpg'].value_counts().to_dict()
    
    # Загружаем нормы и остатки по нужным тпг чтобы посчитать дефицит
    n_link =  r'\\gold585.int\uk\Общее хранилище файлов\Коммерческий департамент\Отдел закупки\ЛИЧНЫЕ\Семен\нужное\матрица обручей.xlsx'
    
    norms = pd.read_excel(n_link, sheet_name='гладкая_плоская размеры_тпг', skiprows=3, usecols=['номер', *tpg]) 
    stock = pd.read_excel(r'C:\Остатки\Обручальные кольца.xlsx', usecols=['КодСклада','Товарная подгруппа','Размер'], dtype='str')
    
    stock['Размер'] = stock['Размер'].str.replace('.', ',')
    stock['tpg'] = stock['Товарная подгруппа'].str.cat(stock['Размер'], sep="_")
    stock.drop(columns=['Товарная подгруппа','Размер'], inplace=True)
    stock = stock[stock['tpg'].isin(tpg)]
    stock = stock.groupby(by=['КодСклада', 'tpg'])['tpg'].aggregate('count').unstack(fill_value=0.0)
    stock.reset_index(inplace=True)
    stock = stock.astype('int')
    stock.columns.values[0] = 'номер'
    stock = stock[norms.columns.tolist()]
    norms = shops.merge(norms, on='номер', how='left')
    stock = shops.merge(stock, on='номер', how='left').fillna(0)
    norms.set_index('номер', inplace=True)
    stock.set_index('номер', inplace=True)
    
    demand = norms - stock
    demand = demand.applymap(lambda num: 0 if num < 0 else num)
    
    def raspred_dict(df):
        emp = []
        cols = df.columns.tolist()
        for col in cols:
            emp.append(list(np.repeat(df.index.to_numpy(),df[col].values.astype(int))))
        return dict(zip(cols, emp))
    
    no7171_dict = raspred_dict(demand)
    
    def dictu(x):
        
        if tpg_dic[x] > len(no7171_dict[x]):
            over = tpg_dic[x] - len(no7171_dict[x])
            no7171_dict[x] = [*no7171_dict[x], *([7171] * over)]
            
        elif tpg_dic[x] < len(no7171_dict[x]):
            no7171_dict[x] = no7171_dict[x][:tpg_dic[x]]
            
        return no7171_dict[x]
    
    lst = [dictu(x) for x in tpg]
    
    prih['Период реализации'] = np.array([item for sublist in lst for item in sublist])
    
    if var_cb1.get() == True:
       prih['Период реализации'] = prih['Период реализации'].apply(lambda x: 7171 if x == 0 else x)
    
    if var_cb2.get() == True:
        prih = prih[prih['Период реализации'] > 0]
    
    name = 'C:/рушники/' + str(random.randrange(5000,10000)) + '.xlsx'
    prih.to_excel(name, sheet_name='Движение товара', index=False) 
    
def distr_bz(file):
    # get stockdata, goods receipt, shops list, stock capability
    prih = pd.read_excel(file)
    prih.sort_values(by=['Товар'], inplace=True)
    prih['Описание 3'] = prih['Описание 3'].str.lower()
    
    letters = prih['Описание 3'].unique().tolist()
    tovs = prih['Товар'].unique().tolist()
    tovs.sort()
    
    stock = pd.read_excel('C:/Остатки/Буквы и Зодиаки.xlsx')
    stock = stock[stock['ТоварНо'].isin(tovs)]
    stock['Описание 3'] = stock['Описание 3'].str.lower()
       
    #estimate sheetname and filename for demand
    tn = prih.loc[1, 'Товарное направление'].lower()
    fname = 'C:/Остатки/' + prih.loc[1, 'Дизайн'][3:].lower() + '.xlsx'
    
    demand = pd.read_excel(fname, sheet_name=tn, skiprows=1, usecols=['Наполняха', 'Код', *letters], index_col='Код')
      
    #Create two dictionaries tov dict and letter dict
    stock_dict = dict(zip(tovs,[stock[stock['ТоварНо'] == tov]['КодСклада'].tolist() for tov in tovs]))
    df = prih.groupby(by=[ 'Товар', 'Описание 3'])['Товар'].count().rename('count')
    df = df.reset_index()
    
    tov_dict = dict(zip(df.Товар.tolist(), df[['Описание 3', 'count']].values.tolist())) #(tov:[letter, demand])
    
    #letter: dataframe with columns:store number, filling rate, demand(positive num is overstock, negative num is demand)
    letter_dict = dict(zip(letters,[demand[['Наполняха', letter]] for letter in letters]))
    if var_cb1.get() == True:
       prih['Период реализации'] = prih['Период реализации'].apply(lambda x: 7171 if x == 0 else x)
    
    if var_cb2.get() == True:
        prih = prih[prih['Период реализации'] > 0]
    
    name = 'C:/рушники/' + str(random.randrange(5000,10000)) + '.xlsx'
    prih.to_excel(name, sheet_name='Движение товара', index=False) 
    
    def get_shoplist(tov):
        """Get sorted by letter demand and stock list of stores with zero stock of tov.
        Changes letter dict according to this list. Two dictionaries needed to perform func"""
        
        letter = tov_dict[tov][0]
        lenth = tov_dict[tov][1]
        shoplist_full = letter_dict[letter].sort_values(by=[letter, 'Наполняха']).index.tolist()
        shoplist = [i for i in shoplist_full if i not in stock_dict[tov]]
        shop_len = len(shoplist)
        
        if shop_len > lenth:
            shoplist = shoplist[:lenth]
        elif len(shoplist) < lenth:
            over = lenth - shop_len
            shoplist = [*shoplist,*([7171] * over)]
        for shop in shoplist:
            if shop != 7171:
                letter_dict[letter][letter][shop] += 1
        return shoplist
          
    lst = [get_shoplist(x) for x in tovs]
    prih['Период реализации'] = np.array([item for sublist in lst for item in sublist])
    
    
def distrfunc():
    path = 'C:/рушники/' + distr_t2.get(1.0, 'end-1c') + '.xlsx'
    
    if var_distr.get() == 0:
        distr_obr(path)
    elif var_distr.get() == 1:
        distr_bz(path)
    else:
        distr_all(path)

#button
distr_button = Button(distr, text="распределить товар", command=distrfunc) 
distr_button.pack(pady=15)

#ORDERS FRAME
orders = Frame(n_book, width=800, height=600)
orders.pack(fill='both', expand=True)

def if_ring():
    # Загружаем таблицу с приходами по всем ТГ
    cs_prih = pd.read_excel(
        r'\\gold585.int\uk\Общее хранилище файлов\Коммерческий департамент\Отдел закупки\ЛИЧНЫЕ\Семен\заказы\4001.xlsx',
        usecols=[0,4,9,11,12,15],names=['Производитель','Тип 1','Номер','Приходы_4001','ТПГ','ОПТ'], skiprows=4)
    
    # Убираем "итог", оптовый товар и нулевые приходы
    
    cs_prih.drop(index=len(cs_prih.Производитель.tolist())-1,inplace=True)
    
    cs_prih = cs_prih[(cs_prih['ОПТ'] == 'Нет') & (cs_prih['Приходы_4001'] != 0)]
    del cs_prih['ОПТ']
    
    # Создаем 2 датафрейма с приходами ИФ и БК
    prih_if = cs_prih[cs_prih['ТПГ']< 1499]
    
    prih_bk = cs_prih[(cs_prih['ТПГ'] > 1499) &
                      (cs_prih['Тип 1'] != 'КОЛЬЦО ОБРУЧАЛЬНОЕ')& 
                      (cs_prih['Тип 1'] != 'СЕРЬГИ-КОНГО')]
    
    del prih_if['ТПГ'], prih_bk['ТПГ'], cs_prih['ТПГ']
    
    # Дефицит и цена ИФ для печаток тоже
    
    import os
    
    def getfiles(dirpath):
        a = [s for s in os.listdir(dirpath) if '~$' not in s]
        a.sort(key=lambda s: os.path.getmtime(os.path.join(dirpath, s)))
        return os.path.join(dirpath, a[-1])
    
    mydir = r'\\gold585.int\uk\Общее хранилище файлов\Коммерческий департамент\Отдел закупки\ЛИЧНЫЕ\Карпова Ксения\ИФ\Дефицит'
    
    deficit_df_if = pd.read_excel(getfiles(mydir),sheet_name='ЛЕНИЗ', 
                                  skiprows=2, usecols=[1,19],names=['tpg','deficit'],nrows=102)
    
    deficit_dict_if = deficit_df_if.set_index('tpg').T.to_dict('records')[0]
    
    price_if = pd.read_excel(
        r'\\gold585.int\uk\Общее хранилище файлов\Коммерческий департамент\Отдел закупки\ЛИЧНЫЕ\Семен\заказы\внав\заказ_и_цены\ИФ.xlsx',
                             usecols=['Номер','цена'], sheet_name='Лист1')
    price_if['цена'] = price_if.цена.round(2)
    
    # Загружаем две таблицы остатки и продажи КОЛЬЦА
    balance_rings = pd.read_excel(
        r'\\gold585.int\uk\Общее хранилище файлов\Коммерческий департамент\Отдел закупки\ЛИЧНЫЕ\Семен\заказы\ифмарт1.xlsx',
        sheet_name='ост_к',skiprows=9, usecols=list(range(9))) #использовать sheet_name='ост' для неколец
    
    sales = pd.read_excel(
        r'\\gold585.int\uk\Общее хранилище файлов\Коммерческий департамент\Отдел закупки\ЛИЧНЫЕ\Семен\заказы\ифмарт1.xlsx',
        sheet_name='прод',skiprows=12,usecols=list(range(2,9)))
    
    # Фильтр приходов по кольцам
    prih_if_rings = prih_if[
        (prih_if['Тип 1'] == 'КОЛЬЦО')|
        (prih_if['Тип 1'] == 'КОЛЬЦО ОБРУЧАЛЬНОЕ')|
        (prih_if['Тип 1'] == 'КОЛЬЦО ПЕЧАТКА')]
        
    # Притягиваем к остаткам продажи, потом добавляем к этому приходы через 'outer'
    
    merged = pd.merge(prih_if_rings,(pd.merge(balance_rings,sales,on='Номер',how = 'left',suffixes=('_ос', '_пр'),  validate='one_to_one')),how = 'outer',  validate='one_to_one')
    
    # соединяем приходы, остатки и продажи
    #merged = pd.merge(prih_if,merged,how = 'outer')
    
    # Добавляем в финальную таблицу недостающую инфу по карточкам
    items = pd.read_excel(
        r'\\gold585.int\uk\Общее хранилище файлов\Коммерческий департамент\Отдел закупки\ЛИЧНЫЕ\Семен\заказы\ифмарт1.xlsx',
        sheet_name='карточки',skiprows=6)
                      
    del items['Описание 3']
    
    final = pd.merge(merged,items,on='Номер',how = 'left',  validate='one_to_one')
    
    # Удаляем опт по группе наценки, ИМ по артикулу
    final = final[~(final['Группа наценки'].str.contains('опт',case=False,na=False))]#na=False
    final = final[~(final['Артикул товара'].str.contains('#ИМ',case=True,na=False))]#na=False
    
    #Суммируем приходы, продажи остатки чтобы и убираем товы с нулями
    #final['сумма']=final.groupby(by='Артикул товара')[col_names[5:7]].sum(axis=1)
    
    # Добавляем дефицит по тпг и цену
    final['Дефицит по тпг'] = final['Товарная подгруппа'].map(deficit_dict_if)
    final = final.merge(price_if,on='Номер',how = 'left', validate='one_to_one')
    
    # Выстраиваем столбцы в нужном порядке и сохраняем
    col_names = final.columns.tolist()
    
    columns = (['Артикул товара', 'Тип 1', 'Производитель', 'Товарная подгруппа'] +
               ['Ценовая корзина', 'Номер', 'Описание', 'Дизайн', 'Размер','Количество камней'] +
               col_names[10:16] + col_names[4:10] +
               ['Приходы_4001', 'Средний вес изделия'] +
               ['Дефицит по тпг','Тип 3','Вставка камней', 'Фото изделия','цена'])
    
    final = final[columns]
    #final=final.set_index('Артикул товара',drop=True, inplace=True)
    
    final.to_excel(
        r'\\gold585.int\uk\Общее хранилище файлов\Коммерческий департамент\Отдел закупки\ЛИЧНЫЕ\Семен\заказы\иф_кольца.xlsx', 
        index=False)
    
def if_other():
    # Загружаем таблицу с приходами по всем ТГ
    cs_prih = pd.read_excel(
        r'\\gold585.int\uk\Общее хранилище файлов\Коммерческий департамент\Отдел закупки\ЛИЧНЫЕ\Семен\заказы\4001.xlsx',
        usecols=[0,4,9,11,12,15],names=['Производитель','Тип 1','Номер','Приходы_4001','ТПГ','ОПТ'], skiprows=4)
    
    # Убираем "итог", оптовый товар и нулевые приходы
    
    cs_prih.drop(index=len(cs_prih.Производитель.tolist())-1,inplace=True)
    
    cs_prih = cs_prih[(cs_prih['ОПТ'] == 'Нет') & (cs_prih['Приходы_4001'] != 0)]
    del cs_prih['ОПТ']
    
    # Создаем 2 датафрейма с приходами ИФ и БК
    prih_if = cs_prih[cs_prih['ТПГ']< 1499]
    
    prih_bk = cs_prih[(cs_prih['ТПГ'] > 1499) &
                      (cs_prih['Тип 1'] != 'КОЛЬЦО ОБРУЧАЛЬНОЕ')& 
                      (cs_prih['Тип 1'] != 'СЕРЬГИ-КОНГО')]
    
    del prih_if['ТПГ'], prih_bk['ТПГ'], cs_prih['ТПГ']
    
    # Дефицит и цена ИФ для печаток тоже
    
    import os
    
    def getfiles(dirpath):
        a = [s for s in os.listdir(dirpath) if '~$' not in s]
        a.sort(key=lambda s: os.path.getmtime(os.path.join(dirpath, s)))
        return os.path.join(dirpath, a[-1])
    
    mydir = r'\\gold585.int\uk\Общее хранилище файлов\Коммерческий департамент\Отдел закупки\ЛИЧНЫЕ\Карпова Ксения\ИФ\Дефицит'
    
    deficit_df_if = pd.read_excel(getfiles(mydir),sheet_name='ЛЕНИЗ', 
                                  skiprows=2, usecols=[1,19],names=['tpg','deficit'],nrows=102)
    
    deficit_dict_if = deficit_df_if.set_index('tpg').T.to_dict('records')[0]
    
    price_if = pd.read_excel(
        r'\\gold585.int\uk\Общее хранилище файлов\Коммерческий департамент\Отдел закупки\ЛИЧНЫЕ\Семен\заказы\внав\заказ_и_цены\ИФ.xlsx',
                             usecols=['Номер','цена'], sheet_name='Лист1')
    price_if['цена'] = price_if.цена.round(2)
    # Загружаем две таблицы остатки и продажи НЕКОЛЬЦА
    balance = pd.read_excel(
        r'\\gold585.int\uk\Общее хранилище файлов\Коммерческий департамент\Отдел закупки\ЛИЧНЫЕ\Семен\заказы\ифмарт1.xlsx',
        sheet_name='ост',skiprows=9) #использовать sheet_name='ост' для неколец
    
    sales = pd.read_excel(
        r'\\gold585.int\uk\Общее хранилище файлов\Коммерческий департамент\Отдел закупки\ЛИЧНЫЕ\Семен\заказы\ифмарт1.xlsx',
        sheet_name='прод',skiprows=12,usecols=list(range(2,9)))
    
    # фильтр приходов по кольцам
    prih_if_norings = prih_if[
        (prih_if['Тип 1'] != 'КОЛЬЦО')&
        (prih_if['Тип 1'] != 'КОЛЬЦО ОБРУЧАЛЬНОЕ')&
        (prih_if['Тип 1'] != 'КОЛЬЦО ПЕЧАТКА')
    ]
    
    # Притягиваем к остаткам продажи, потом добавляем к этому приходы через 'outer'
    merged = pd.merge(prih_if_norings,(pd.merge(balance,sales,on='Номер',how = 'left',suffixes=('_ос', '_пр'))),how = 'outer')
    
    # добавляем в финальную таблицу недостающую инфу по карточкам
    items = pd.read_excel(
        r'\\gold585.int\uk\Общее хранилище файлов\Коммерческий департамент\Отдел закупки\ЛИЧНЫЕ\Семен\заказы\ифмарт1.xlsx',
        sheet_name='карточки',skiprows=6)
    del items['Размер']
    
    final = pd.merge(merged,items,on='Номер',how = 'left')
    
    # Удаляем опт, ИМ
    final=final[~(final['Группа наценки'].str.contains('опт',case=False,na=False))]
    final=final[~(final['Артикул товара'].str.contains('#ИМ',case=True,na=False))]
    
    # Добавляем дефицит по тпг и цену
    final['Дефицит по тпг'] = final['Товарная подгруппа'].map(deficit_dict_if)
    final = final.merge(price_if,on='Номер',how = 'left', validate='one_to_one')
    
    # Выстраиваем столбцы в нужном порядке
    col_names = final.columns.tolist()
    
    columns = (['Артикул товара', 'Тип 1', 'Производитель', 'Товарная подгруппа'] +
               ['Ценовая корзина', 'Номер', 'Описание', 'Дизайн', 'Описание 3','Количество камней'] +
               col_names[10:16] + col_names[4:10] +
               ['Приходы_4001', 'Средний вес изделия'] +
               ['Дефицит по тпг', 'Тип 3','Вставка камней', 'Фото изделия', 'цена'])
    
    final = final[columns]
    
    #Суммируем приходы, продажи остатки чтобы и убираем товы с нулями
    #final['сумма'] = final[col_names[3:18]].sum(axis=1)
    #final=final[final['сумма'] >0]
    #del final['сумма']
    
    #Удаляем буквы и зодиаки
    final = final[(final['Дизайн'] !='ИФ БУКВЫ') & (final['Дизайн'] !='ИФ ЗОДИАК')]
    
    final.to_excel(
        r'\\gold585.int\uk\Общее хранилище файлов\Коммерческий департамент\Отдел закупки\ЛИЧНЫЕ\Семен\заказы\иф_некольца.xlsx', 
        index=False)

def bk():
    # Загружаем таблицу с приходами по всем ТГ
    cs_prih = pd.read_excel(
        r'\\gold585.int\uk\Общее хранилище файлов\Коммерческий департамент\Отдел закупки\ЛИЧНЫЕ\Семен\заказы\4001.xlsx',
        usecols=[0,4,9,11,12,15],names=['Производитель','Тип 1','Номер','Приходы_4001','ТПГ','ОПТ'], skiprows=4)
    
    # Убираем "итог", оптовый товар и нулевые приходы
    
    cs_prih.drop(index=len(cs_prih.Производитель.tolist())-1,inplace=True)
    
    cs_prih = cs_prih[(cs_prih['ОПТ'] == 'Нет') & (cs_prih['Приходы_4001'] != 0)]
    del cs_prih['ОПТ']
    
    # Создаем 2 датафрейма с приходами ИФ и БК
    prih_if = cs_prih[cs_prih['ТПГ']< 1499]
    
    prih_bk = cs_prih[(cs_prih['ТПГ'] > 1499) &
                      (cs_prih['Тип 1'] != 'КОЛЬЦО ОБРУЧАЛЬНОЕ')& 
                      (cs_prih['Тип 1'] != 'СЕРЬГИ-КОНГО')]
    
    del prih_if['ТПГ'], prih_bk['ТПГ'], cs_prih['ТПГ']
    # Загружаем остатки, продажи и карточки
    balance = pd.read_excel(
        r'\\gold585.int\uk\Общее хранилище файлов\Коммерческий департамент\Отдел закупки\ЛИЧНЫЕ\Семен\заказы\бк.xlsx',
        sheet_name='ост',skiprows=9)
    
    sales = pd.read_excel(
        r'\\gold585.int\uk\Общее хранилище файлов\Коммерческий департамент\Отдел закупки\ЛИЧНЫЕ\Семен\заказы\бк.xlsx',
        sheet_name='прод',skiprows=13, usecols=list(range(2,9)))
    
    items = pd.read_excel(
        r'\\gold585.int\uk\Общее хранилище файлов\Коммерческий департамент\Отдел закупки\ЛИЧНЫЕ\Семен\заказы\бк.xlsx',
        sheet_name='карточки',skiprows=6)
    
    # Соединяем остатки и продажи, добавляем приходы через 'outer', добавляем инфу из карточек
    merged = pd.merge(prih_bk,(pd.merge(balance,sales,on='Номер',how = 'left',suffixes=('_ос', '_пр'), validate='one_to_one')),how = 'outer', validate='one_to_one')
    final = pd.merge(merged,items,on='Номер',how = 'left', validate='one_to_one')
    
    # Удаляем опт, ИМ
    final = final[~(final['Группа наценки'].str.contains('опт',case=False,na=False))]#na=False
    final = final[~(final['Артикул товара'].str.contains('#ИМ',case=True,na=False))]#na=False
    
    # Создаем словарь с дефицитом по тпг и добавляем данные в таблицу
    deficit_df_bk = pd.read_excel(
        r'\\gold585.int\uk\Общее хранилище файлов\Коммерческий департамент\Отдел закупки\ЛИЧНЫЕ\Ксю\Дефицит по ТПГ 2020.xlsx',
        sheet_name='Дефицит', skiprows=2, usecols=[0,18],names=['tpg','deficit'],nrows=97)
    
    deficit_df_bk['tpg'] = deficit_df_bk.tpg.astype('float64')
    deficit_dict_bk = dict(zip(deficit_df_bk.tpg, deficit_df_bk.deficit))
    final['дефицит по тпг'] = final['Товарная подгруппа'].map(deficit_dict_bk)
    
    # Добавляем информацию по цене
    price_bk = pd.read_excel(
        r'\\gold585.int\uk\Общее хранилище файлов\Коммерческий департамент\Отдел закупки\ЛИЧНЫЕ\Назарова А.С.\Цены БК.xlsx',
        sheet_name='Нов', usecols=['Номер','цена']) 
    
    final = final.merge(price_bk, on='Номер', how='left', validate='one_to_one')
    
    # Выстраиваем столбцы в нужном порядке
    ostatki_prodazi = final.columns.tolist()[10:16] + final.columns.tolist()[4:10]
    
    columns = ['Артикул товара', 'Тип 1', 'Производитель', 'Товарная подгруппа', 'Ценовая корзина', 'Номер', 'Описание', 
               'Дизайн', 'Размер', 'Описание 3', *ostatki_prodazi, 'Приходы_4001','Средний вес изделия', 'дефицит по тпг', 
               'Тип 3', 'Фото изделия', 'Цена золото', 'Тип изготовления']
    
    final = final[columns]
    
    #final=final.set_index('Артикул товара',drop=True, inplace=True)
    
    # Фильтруем по кольцам убираем описание 3 и суммируем приходы 4001+4093 в "приходы". 
    final_ring = final[final['Тип 1'] =='КОЛЬЦО'].drop(['Описание 3', 'Тип изготовления'], axis=1)
    
    
    # Фильтруем по некольцам и убираем размер и буквы с зодиаками
    final_other = final[~final['Тип 1'].str.contains('КОЛЬЦО')].drop(['Размер'], axis=1)
    
    final_other = final_other[~(
        ((final_other['Тип 1'] == 'ПОДВЕС ДЕКОРАТИВНЫЙ') & 
         ((final_other['Дизайн'] == 'ИФ ДВУСПЛАВ ДЕКОР') | 
          (final_other['Дизайн'] =='ИФ БУКВЫ') | 
          (final_other['Дизайн'] =='ИФ ЗОДИАК'))) | 
        (final_other['Тип 1'] == 'СЕРЬГИ-КОНГО') | 
        (final_other['Тип 1'] == 'МОНЕТА')
        
    )]
    
    # сохраняем кольца и некольца на отдельные вкладки
    with pd.ExcelWriter (
         r'\\gold585.int\uk\Общее хранилище файлов\Коммерческий департамент\Отдел закупки\ЛИЧНЫЕ\Семен\заказы\заказы_бк.xlsx'
    ) as writer:
        final_ring.to_excel(writer, sheet_name='кольца', index=False)
        final_other.to_excel(writer, sheet_name='некольца', index=False)

def ringstones():
    # Загружаем остатки, продажи, карточки
    path = '//gold585.int/uk/Общее хранилище файлов/Коммерческий департамент/Отдел закупки/ЛИЧНЫЕ/Семен/заказы/печатки.xlsx'
    
    balance = pd.read_excel(path, sheet_name='ост_к',skiprows=9, usecols=list(range(9)))
    sales = pd.read_excel(path, sheet_name='прод',skiprows=12,usecols=list(range(2,9)))
    items = pd.read_excel(path, sheet_name='карточки',skiprows=6)
    
    # Соединяем остатки и продажи, добавляем приходы через 'outer', добавляем инфу из карточек
    merged = pd.merge(cs_prih[cs_prih['Тип 1'] == 'КОЛЬЦО ПЕЧАТКА'],(pd.merge(balance,sales,on='Номер',how = 'left',suffixes=('_ос', '_пр'))),how = 'outer')
    final = pd.merge(merged,items,on='Номер',how = 'left')
    
    # Удаляем опт, ИМ
    final = final[~(final['Группа наценки'].str.contains('опт',case=False,na=False)) & 
                  ~(final['Артикул товара'].str.contains('#ИМ',case=True,na=False))]#na=False
    
    # Создаем таблицу с дефицитом по тпг и добавляем данные в основную таблицу
    deficit_df_bk = pd.read_excel(
        r'\\gold585.int\uk\Общее хранилище файлов\Коммерческий департамент\Отдел закупки\ЛИЧНЫЕ\Ксю\Дефицит по ТПГ 2020.xlsx',
        sheet_name='Дефицит', skiprows=2, usecols=[0,18],names=['tpg','deficit'],nrows=97)
    
    pech_if = deficit_df_if[deficit_df_if['tpg'].isin([1028,1029,1030])]
    pech_bk = deficit_df_bk[deficit_df_bk['tpg'].isin([1538,1539,1540])]
    
    deficit_pech = pd.concat([pech_if, pech_bk], ignore_index=True)
    deficit_pech.rename(columns={'tpg':'Товарная подгруппа', 'deficit':'Дефицит по тпг'},inplace=True)
    
    final = final.merge(deficit_pech, on='Товарная подгруппа',how = 'left')
    
    # Добавляем цены
    price_bk = pd.read_excel(
        r'\\gold585.int\uk\Общее хранилище файлов\Коммерческий департамент\Отдел закупки\ЛИЧНЫЕ\Назарова А.С.\Цены БК.xlsx',
        sheet_name='Нов', usecols=['Номер','цена']) 
    price_bk.rename(columns={'Цена золото':'цена'}, inplace=True)
    
    price_pech = pd.concat([price_bk, price_if], ignore_index=True)
    final = final.merge(price_pech, on='Номер',how = 'left')
    
    # Выстраиваем столбцы в нужном порядке и сохраняем
    col_names = final.columns.tolist()
    
    columns = (
        ['Артикул товара', 'Тип 1', 'Производитель', 'Товарная подгруппа', 'Ценовая корзина', 'Номер', 'Описание', 'Дизайн', 
         'Размер','Количество камней', *col_names[10:16], *col_names[4:10], 'Приходы_4001', 'Средний вес изделия', 
         'Дефицит по тпг','Тип 3','Вставка камней', 'Фото изделия','цена'])
    
    final = final[columns]
    #final=final.set_index('Артикул товара',drop=True, inplace=True)
    
    final.to_excel(
        r'\\gold585.int\uk\Общее хранилище файлов\Коммерческий департамент\Отдел закупки\ЛИЧНЫЕ\Семен\заказы\заказ_печатки.xlsx', 
        index=False)

def demand():
    # Загружаем остатки и нормы
    stock = pd.read_excel(r'C:\Остатки\Обручальные кольца.xlsx', usecols=['КодСклада','Товарная подгруппа','Размер'])
    
    n_link =  r'\\gold585.int\uk\Общее хранилище файлов\Коммерческий департамент\Отдел закупки\ЛИЧНЫЕ\Семен\нужное\матрица обручей.xlsx'
    
    norms = pd.read_excel(n_link, sheet_name='гладкая_плоская размеры_тпг', skiprows=3)
    del norms['адрес']
    del norms['норма']
    
    # выбрать остатки 1517-1524 тпг (без 1521-6,8)
    stock = stock[(stock['Товарная подгруппа'] > 1516) & 
                  (stock['Товарная подгруппа'] < 1525) & 
                  (stock['Товарная подгруппа'] != 1521)]
    
    # убрать нули в тпг и размерах и точку поменять на запятую
    stock['Товарная подгруппа'] = stock['Товарная подгруппа'].astype('int')
    stock['Размер'] = stock['Размер'].apply(lambda size: str(int(size)) if size == int(size) else str(size))
    stock['Размер'] = stock['Размер'].apply(lambda size: size.replace('.',','))
    
    # добавим в остатки колонку тпг_размер, группируем по этой колонке и считаем
    stock[['Товарная подгруппа','Размер']] = stock[['Товарная подгруппа','Размер']].astype('str')
    stock['size_tpg'] = stock['Товарная подгруппа'].str.cat(stock['Размер'], sep='_')
    stock = stock.groupby(['КодСклада','size_tpg'])['size_tpg'].count().unstack().reset_index()
    
    # Берем столбец с магами из норм и соединяем с остатками
    stock.rename(columns={'КодСклада':'номер'},inplace=True)
    stock = norms[['номер']].merge(stock, how='left')
    
    # добавим недостающие колонки в остатки, чтобы получить идентичный нормам датафрейм
    zero_cols = list(set(norms.columns) - set(stock.columns))
    
    for cols in zero_cols:
        stock[cols] = np.nan
        
    stock = stock[norms.columns.to_list()].fillna(0)
    
    # Словарь соответствия тпг-дизайн+толщина
    design_dict = {'1517':'ИФ ОБРУЧ ГЛАДКАЯ2', '1518':'ИФ ОБРУЧ ГЛАДКАЯ3','1519':'ИФ ОБРУЧ ГЛАДКАЯ4','1520':'ИФ ОБРУЧ ГЛАДКАЯ5',
                   '1521':'ИФ ОБРУЧ ГЛАДКАЯ6','1522':'ИФ ОБРУЧ ПЛОСКАЯ3','1523':'ИФ ОБРУЧ ПЛОСКАЯ4','1524':'ИФ ОБРУЧ ПЛОСКАЯ5'}

    # Вычитаем один датафрейм из другого, отрицательные значения меняем на нули, считаем сумму по строкам и сохраняем
    if all(norms.columns == stock.columns):
        demand = norms - stock
        demand = demand.applymap(lambda num: 0 if num < 0 else num)
        demand = demand.T.iloc[1:]
        demand['sum'] = demand.sum(axis=1)
        demand.reset_index(inplace=True)
        demand['index'] = demand['index'].apply(lambda x: design_dict[x[:4]] + x[4:])
        demand[['index','sum']].to_excel('C:\Остатки\деф_Обручальные кольца.xlsx', index=False)
    else:
        print('столбцы не совпадают')    
#buttons       
if_ringb = Button(orders, text="Кольца ИФ", command=if_ring)
if_otherb = Button(orders, text="НЕ кольца ИФ", command=if_other)
bkb = Button(orders, text="БК", command=bk)
ringstonesb = Button(orders, text="Печатки ИФБК", command=ringstones)
demandb = Button(orders, text="Дефицит по обручальным кольцам", command=demand)

btnlst = [if_ringb, if_otherb, bkb, ringstonesb, demandb] 
for x in btnlst:
    x.pack(pady=7)
    x.configure(width=100)
#PROCESS FRAME
process = Frame(n_book, width=800, height=600)
process.pack(fill='both', expand=True)

# blocks
fr1 = Frame(process)
fr2 = Frame(process)
fr3 = Frame(process)

for x in [fr1, fr2, fr3]:
    x.pack(ipady = 15)
    
lbl1 = Label(fr1,text='Создать списки накладных')
lbl2 = Label(fr2,text='Порезать накладную')
lbl3 = Label(fr3,text='Форматировать накладную')

for x in [lbl1, lbl2, lbl3]:
    x.pack(pady=5)

#block1
varpr_cb1 = BooleanVar()
varpr_cb1.set(False)
pr_cb1 = Checkbutton(fr1, text='ИФБК отдельно', variable=varpr_cb1)
pr_cb1.pack()

pr_txt1 = Text(fr1, width=600, height=10)
pr_txt1.pack()

def prih_to_Nav():
    path = 'C:/анаконда/Book1.xlsx'
    df = pd.read_excel(path,usecols=[3, 10, 14, 15, 17], skiprows=4, names=['prih','amount','sklad', 'hand', 'tn'])
    df = df[(df.tn.isin(['БК', 'ИФ'])) & (df.sklad == 'Центральный склад ЛенИЗ') & (df.hand != 'Да')]
    df.drop(columns=['sklad', 'hand'], inplace=True)
    df_if = df[df.tn == 'ИФ']
    df_bk = df[df.tn == 'БК']
    def create_line(list):
        return "".join(["|" + item if item != list[0] else list[0] for item in list])
    
    def cut_prihs(dataframe):
        # делим на группы по 3000
        bins = int(np.ceil(dataframe.amount.sum()/3500))
        label = list(range(0, bins))
        dataframe['group'] = pd.cut(dataframe.amount.cumsum(), bins=bins, labels=label)
        prih_lst = [create_line(dataframe.prih[dataframe.group == x].tolist()) for x in label]
        return prih_lst
    
    if varpr_cb1.get() == True:
       prih_lst = ['ИФ'] + cut_prihs(df_if) + ['БК'] + cut_prihs(df_bk)
    else:
        prih_lst = cut_prihs(df)
        
    pr_txt1.insert(1.0, np.array(prih_lst))

pr_b1 = Button(fr1, text="Создать", command=prih_to_Nav)
pr_b1.pack(pady=7)

#block2
pr_txt2 = Text(fr2, width=8, height=1)
pr_txt2.pack()

varpr_cmb = StringVar()
cmb_pr = Combobox(fr2, width = 27, textvariable = varpr_cmb)

cmb_pr['values'] = (
        'Дизайн', 'Товар', 'Тип изделия 1', 'Тип изделия 3', 
        'Артикул', 'Товарное направление', 'Дата расхода'
        )
cmb_pr.pack()

def cutbycolumn():
    folder = 'C:/рушники/'
    ending = '.xlsx'
    path = folder + pr_txt2.get(1.0, "end-1c") + ending
    df = pd.read_excel(path)
    column = varpr_cmb.get()
    lst = df[column].unique().tolist()
    for x in lst:
        df[df[column] == x].to_excel(folder + str(x) + ending,
          sheet_name='Движение товара', index=False)

pr_b2 = Button(fr2, text="Порезать", command=cutbycolumn)
pr_b2.pack(pady=7) 

#block3
varpr_cb2 = BooleanVar()
varpr_cb2.set(False)
pr_cb2 = Checkbutton(fr3, text='Выбрать товы', variable=varpr_cb2)
pr_cb2.pack()

pr_txt3 = Entry(fr3, width=500)
pr_txt3.pack(pady=5)   

def create_zab():
    prih = pd.read_excel(r'C:\Users\Dotsenko.Semen\Documents\Книга1.xlsx')
    del prih['№ Поставщика']
    if varpr_cb2.get() == True:
        tov_lst = str(pr_txt3.get()).split(',')
        prih = prih[prih['Товар'].isin(tov_lst)]

    name = 'C:/рушники/' + str(random.randrange(5000,10000)) + '.xlsx'
    prih.to_excel(name, sheet_name='Движение товара', index=False)
    pr_txt4.insert(1.0, name[11:15])
    
pr_b3 = Button(fr3, text="Преобразовать", command=create_zab)
pr_b3.pack(pady=7) 

pr_txt4 = Text(fr3, width=8, height=1)
pr_txt4.pack(side=BOTTOM)


    
#NORMS FRAME
norms = Frame(n_book, width=500, height=500)
norms.pack(fill='both', expand=True)

var_norms =  IntVar()
norms_dict = {"Нормы по понедельникам": 0, 
              "Нормы по буквам и зодиакам": 1,
              "Список акутальных магов с нормами по ТГ": 2, 
              "Нормы по ТГ для рейтинга по ТН": 3}
 
for (text, value) in norms_dict.items():
    Radiobutton(norms, text = text, variable = var_norms,
        value = value).pack(side = TOP, ipady = 5)

def monday_norms():
    # Загружаем нужные столбцы
    cols = ['Номер из NAV','Адрес магазина','дата планового открытия','статус','ИФ кольца','ИФ КОЛЬЦА ОБРУЧ',
            'ИФ печатки','ИФ серьги','ИФ подвес культ','ИФ подвес декор','БК кольца','БК кольца обруч','БК печатки',
            'БК серьги','БК подвес культ','БК подвес декор','ИТОГО, нормы в штуках, без накоплений']
    
    df = pd.read_excel(
        r'\\gold585.int\uk\Общее хранилище файлов\Коммерческий департамент\реестр магазинов\РЕЕСТР МАГАЗИНОВ NEW.xlsm',
        sheet_name='БАЗА',skiprows=2, usecols=cols)
    
    # Убираем ненужную инфу
    df.rename({'ИТОГО, нормы в штуках, без накоплений':"sum_all"}, axis=1, inplace=True)
    df = df[df['sum_all'] > 0]
    df['статус'] = df.статус.str.lower()
    df = df[df.статус.isin(['открыт','склад','отгружен'])]
    
    # Добавляем сумму по ИФ\БК и оставляем этот столбец вместо столбцов с нормами по иф\бк
    sum_cols = [x for x in cols if 'ИФ' in x or 'БК' in x ]
    df['sum_gold'] = df[sum_cols].sum(axis=1)
    df.drop(axis=1, columns=sum_cols, inplace=True)
    
    # Количество открытых и к открытию всего и только по золоту
    shops_count = pd.DataFrame({
        'Магов открыто': [df.sum_all[df.статус !='склад'].count(), df.sum_gold[(df.sum_gold > 0) & (df.статус !='склад')].count()],
        'Магов к открытию': [df.sum_all[df.статус =='склад'].count(), df.sum_gold[(df.sum_gold > 0) & (df.статус =='склад')].count()],
        'Итого магов': [df.sum_all.count(),df.sum_gold[df.sum_gold > 0].count()]})
    
    # Список новых магов с датами
    new_shops = df[['Номер из NAV', 'Адрес магазина', 'дата планового открытия']][(df.sum_gold > 0) & (df.статус =='склад')]
    new_shops.sort_values('дата планового открытия', inplace=True)
    
    # Нормы ИФ/БК по ТПГ
    row_number = pd.read_excel(
        r'\\gold585.int\uk\Общее хранилище файлов\Коммерческий департамент\реестр магазинов\РЕЕСТР МАГАЗИНОВ NEW.xlsm',
        sheet_name='НОРМЫ 17.12',skiprows=4,usecols=[5])
    
    skip_row = row_number.ЦОК.tolist()
    skip_row_n = skip_row.index('подгруппа')-1
    
    norm_all = pd.read_excel(
        r'\\gold585.int\uk\Общее хранилище файлов\Коммерческий департамент\реестр магазинов\РЕЕСТР МАГАЗИНОВ NEW.xlsm',
        sheet_name='НОРМЫ 17.12',skiprows=4,nrows=skip_row_n)
    
    lst = norm_all.columns.tolist()
    
    cols = (
        [lst[lst.index('Название из NAV')]] +
        lst[lst.index(1001):lst.index(1102)+1] +
        lst[lst.index(1501):lst.index(1597)+1]
    )
    norm_all = norm_all[cols]
    gold_shops = df[['Номер из NAV', 'статус']][df.sum_gold > 0]
    merged = pd.merge(gold_shops, norm_all, left_on='Номер из NAV', right_on='Название из NAV', how = 'left')
    
    del merged['Название из NAV'], merged['Номер из NAV']
    
    # Все нецифры меняю на ноль. из ексель подтягивается значок '-'
    nostatus = [i for i in merged.columns.to_list() if i !='статус']
    merged[nostatus] = merged[nostatus].apply(pd.to_numeric, errors='coerce').fillna(0)
    
    
    merged = merged.groupby(by='статус').sum()
    merged = merged.T
    merged['sum_all'] = merged.sum(axis=1)
    merged = merged[['открыт', 'sum_all']]
    
    # СОХРАНЯЕМ ВСЕ В 1 ФАЙЛ
    
    with pd.ExcelWriter (
        r'\\gold585.int\uk\Общее хранилище файлов\Коммерческий департамент\Отдел закупки\ЛИЧНЫЕ\Семен\нужное\для_новые_нормы.xlsx'
    ) as writer:
        shops_count.to_excel(writer, sheet_name='количество', index=False)
        new_shops.to_excel(writer, sheet_name='новые_маги', index=False)
        merged.to_excel(writer, sheet_name='новые_нормы')
        
def bz_norms():
    cols = [2, *range(368,374), *range(469,475)]
    link = r'\\gold585.int\uk\Общее хранилище файлов\Коммерческий департамент\реестр магазинов\РЕЕСТР МАГАЗИНОВ NEW.xlsm'    
    norm = pd.read_excel(link, sheet_name='НОРМЫ 17.12', skiprows=4, usecols=cols)
    
    # Обрезаем пустые строки, делим на буквы и зодиаки
    last_row = norm.index[norm['Название из NAV'].isna()][0]-1
    norm = norm.loc[:last_row]
    norm.set_index('Название из NAV',inplace=True)
    
    b_columns = [1088, 1089, 1090, 1586, 1587, 1588]
    z_columns = [1091, 1092, 1093, 1589, 1590, 1591]
    norm_b = norm[b_columns]
    norm_z = norm[z_columns]
    
    norm_b = norm_b.apply(pd.to_numeric, errors='coerce').fillna(0)
    norm_z = norm_z.apply(pd.to_numeric, errors='coerce').fillna(0)
    
    norm_b['sum_if'], norm_b['sum_bk'] = norm_b[[1088, 1089, 1090]].sum(axis=1), norm_b[[1586, 1587, 1588]].sum(axis=1)
    norm_z['sum_if'], norm_z['sum_bk'] = norm_z[[1091, 1092, 1093]].sum(axis=1), norm_z[[1589, 1590, 1591]].sum(axis=1)
    
    with pd.ExcelWriter (r'C:\Остатки\norm_bz.xlsx') as writer:
        norm_b.to_excel(writer, sheet_name='norm_b')
        norm_z.to_excel(writer, sheet_name='norm_z')
    
def operating_shops():
    cols_df = ['Номер из NAV', 'Адрес магазина', 'дата планового открытия','статус', 'ИФ кольца', 'ИФ КОЛЬЦА ОБРУЧ', 'ИФ печатки', 'ИФ серьги', 
            'ИФ подвес культ', 'ИФ подвес декор', 'БК кольца', 'БК кольца обруч', 'БК печатки', 'БК серьги', 
            'БК подвес культ', 'БК подвес декор']
          
    df = pd.read_excel(
        r'\\gold585.int\uk\Общее хранилище файлов\Коммерческий департамент\реестр магазинов\РЕЕСТР МАГАЗИНОВ NEW.xlsm',
        sheet_name='БАЗА',skiprows=2, usecols=cols_df)
    
    #выбираем открытые маги с суммой норм по иф/бк больше нуля
    df = df[df.статус.str.contains('откр|отгр|склад', na=False, case=False)]
    df['сумма'] = df.iloc[:, 3:].sum(axis=1)
    df = df[df['сумма'] > 0]
    df = df[df['Номер из NAV'] != 3323]# Тамбов носовская
        
    all_shops = df.drop(columns=['статус','сумма','дата планового открытия'])
    all_shops = all_shops.rename(columns={'Номер из NAV':'Код'})
    all_shops['Код'] = all_shops['Код'].astype(np.int64)
    
    # делим на два отдельных иф и бк
    cols = all_shops.columns.tolist()
    shops_if = all_shops[[*cols[:5],cols[7],cols[6],cols[5]]]
    shops_bk = all_shops[[*cols[:2], *cols[8:11], cols[13],cols[12],cols[11]]]
    
    #выгружаем из файла "наполняемость" актуальную структуру магов иф/бк, добавляет туда нормы и новые маги из датафрэйма выше
    shops_if_list = pd.read_excel(
        r'C:\Остатки\наполняемость с 10012020.xlsx',
        sheet_name='ИФ',skiprows=3, usecols=[2],keep_default_na=False)
    
    shops_bk_list = pd.read_excel(
        r'C:\Остатки\наполняемость с 10012020.xlsx',
        sheet_name='БК',skiprows=3, usecols=[2],keep_default_na=False)
    
    # список новых магов которые надо добавить
    add_to_if = pd.DataFrame({'Код':[x for x in shops_if['Код'].tolist() if x not in shops_if_list['Код'].tolist()]})
    add_to_bk = pd.DataFrame({'Код':[x for x in shops_bk['Код'].tolist() if x not in shops_bk_list['Код'].tolist()]})
    
    # добавляем новые маги в структуру
    shops_if_list = shops_if_list.append(add_to_if, sort=False)
    shops_bk_list = shops_bk_list.append(add_to_bk, sort=False)
    
    #соединяем "структуру" и нормы
    shops_if = pd.merge(shops_if_list,shops_if,how='left',on='Код')
    shops_bk = pd.merge(shops_bk_list,shops_bk,how='left',on='Код')
        
    #удаляем индекс и сохраняем на отдельные вкладки
    shops_if.set_index('Код', drop=True,inplace=True)
    shops_bk.set_index('Код', drop=True,inplace=True)
    
    #создаем датафрейм из магов с датами открытия последние 90 дней
    from datetime import date, timedelta
    
    three_month_ago = date.today() - timedelta(90)
    df['дата планового открытия'] = df['дата планового открытия'].apply(lambda x:x.date())#чтобы ошибка не выпадала о формате
    dates = df[['Номер из NAV', 'дата планового открытия']][df['дата планового открытия'] > three_month_ago]
    dates.set_index('Номер из NAV', drop=True,inplace=True)
    
    #сохраняем все в эксель
    with pd.ExcelWriter (
        r'\\gold585.int\uk\Общее хранилище файлов\Коммерческий департамент\Отдел закупки\ЛИЧНЫЕ\Семен\актуальные маги.xlsx'
    ) as writer:
        shops_if.to_excel(writer, sheet_name='ИФ')
        shops_bk.to_excel(writer, sheet_name='БК')
        dates.to_excel(writer, sheet_name='даты')

def normsforraiting():
    #читает файл с реестром
    df = pd.read_excel(
        r'\\gold585.int\uk\Общее хранилище файлов\Коммерческий департамент\реестр магазинов\РЕЕСТР МАГАЗИНОВ NEW.xlsm',
        sheet_name='БАЗА',skiprows=2)
    
    #вынимает нужные две строки и сохраняет
    rep_tn = df.iloc[0:1,28:74]
    rep_tn.to_excel(
        r'\\gold585.int\uk\Общее хранилище файлов\Коммерческий департамент\Отдел закупки\ЛИЧНЫЕ\Татьяна ИФ.БК\РЕЙТИНГИ\Рейтинг ТГ и ТН\123.xlsx',
    index=False)

def create_norms():
    func_dict = {0: monday_norms(), 1: bz_norms(), 2: operating_shops(), 3: normsforraiting()}
    return func_dict[var_norms.get()]
    
norms_b1 = Button(norms, text="Сформировать файл с нормами", command=create_norms)
norms_b1.pack(pady=10)   

var_norms2 =  IntVar()
norms_dict = {"ТОП_100": 0, 
              "Обручальные кольца": 1,
              "Мусульманские подвесы": 2, 
              "Приоритеты по ТГ": 3}
 
for (text, value) in norms_dict.items():
    Radiobutton(norms, text = text, variable = var_norms2,
        value = value).pack(side = TOP, ipady = 5)
    
def table_tonav(df):
    """Превращает таблицу, где первый столбец маги, а остальные нормы по тпг в таблицу для загрузки этих норм в нав. 
    По умолчанию максимальное значение 1, коэф отставания 0 индекс = номер магазина"""    
    
    norms_list = df.T.to_numpy().tolist()
    
    tpg_list = df.columns.tolist()
    
#добавим бк/иф к названиям тпг   
    new_tpg_list = ['БК_' + str(tpg) if int(str(tpg)[:4]) >= 1500 else 'ИФ_' + str(tpg) for tpg in tpg_list]
          
    tpg = [x for x in new_tpg_list for i in range(len(df.index.tolist()))]
    shop = df.index.tolist()*len(df.columns.tolist())
    norms = [round(item) for sublist in norms_list for item in sublist]#Все нормы округляются
    data = {'Товарная Группа':tpg,'Норма':norms,'Макс. допустимое кол-во на точке':[1]*len(norms),'коэф. Корр отставания проведения продаж':[0]*len(norms)}
    new_df = pd.DataFrame(index=shop, data=data)
    new_df.index.name = 'Код'                    
    return new_df

def priority_tonav(df):
    
#индекс = номер магазина    
    norms_list = df.T.to_numpy().tolist()
    tpg_list = df.columns.tolist()
    tpg = [x for x in tpg_list for i in range(len(df.index.tolist()))]
    shop = df.index.tolist()*len(df.columns.tolist())
    norms = [item for sublist in norms_list for item in sublist]#Все нормы округляются
    data = {'Товарное направление':tpg,'Приоритет':[100]*len(norms),'Повышающий коэффициент':[1]*len(norms),'Исключить из распределения':norms}
    new_df = pd.DataFrame(index=shop, data=data)
    new_df.index.name = 'Код'                    
    return new_df

def topnorms():
    bk_df = pd.read_excel(r'\\gold585.int\uk\Общее хранилище файлов\Коммерческий департамент\Отдел закупки\ЛИЧНЫЕ\Татьяна ИФ.БК\РЕЙТИНГИ\топ100.xlsx', sheet_name='тпг_бк', index_col=1, skiprows=1)
    if_df = pd.read_excel(r'\\gold585.int\uk\Общее хранилище файлов\Коммерческий департамент\Отдел закупки\ЛИЧНЫЕ\Татьяна ИФ.БК\РЕЙТИНГИ\топ100.xlsx', sheet_name='тпг_иф', index_col=1, skiprows=1)
    
    del bk_df['ранг']
    del if_df['ранг']
    
    result = table_tonav(bk_df).append(table_tonav(if_df))
    
    result.to_excel(r'\\gold585.int\uk\Общее хранилище файлов\Коммерческий департамент\Отдел закупки\ЛИЧНЫЕ\Семен\в_нав.xlsx') 

def wedring_norms():
    #загружаем нормы по размерам и по тпг
    wed_rings1 = pd.read_excel(
        r'\\gold585.int\uk\Общее хранилище файлов\Коммерческий департамент\Отдел закупки\ЛИЧНЫЕ\Семен\нужное\матрица обручей.xlsx',
        sheet_name='гладкая_плоская размеры_тпг',index_col=0,skiprows=3)
    del wed_rings1['адрес']
    del wed_rings1['норма']
    
    ring_cols = [0, 3, 8] + list(range(12,25))
    wed_rings2 = pd.read_excel(
        r'\\gold585.int\uk\Общее хранилище файлов\Коммерческий департамент\Отдел закупки\ЛИЧНЫЕ\Семен\нужное\матрица обручей.xlsx',
        sheet_name='Матрица по тпг',index_col=0,skiprows=3, usecols=ring_cols)
    
    #применяем функцию исправляем максимумы соединяем в один датафрейм и сохраняем
    wed_rings_tonav1 = table_tonav(wed_rings1)
    wed_rings_tonav1['Макс. допустимое кол-во на точке'] = wed_rings_tonav1['Норма']
    wed_rings_tonav2 = table_tonav(wed_rings2)
    
    wed_rings_tonav = pd.concat([wed_rings_tonav1, wed_rings_tonav2 ])
    
    wed_rings_tonav.to_excel(
        r'\\gold585.int\uk\Общее хранилище файлов\Коммерческий департамент\Отдел закупки\ЛИЧНЫЕ\Семен\в_нав.xlsx')    
    
def muslim_norms():
    from_26 = [0] + list(range(26,49))
    mus_df = pd.read_excel(
        r'\\gold585.int\uk\Общее хранилище файлов\Коммерческий департамент\Отдел закупки\ЛИЧНЫЕ\Семен\нужное\мусульманкэ.xlsx',
        sheet_name='нормы культ',index_col=0,skiprows=2,usecols=from_26)
    
    mus_df.columns = [x[:4] for x in mus_df.columns.tolist()]
    
    mus_df = table_tonav(mus_df)
    
    mus_df.to_excel(
        r'\\gold585.int\uk\Общее хранилище файлов\Коммерческий департамент\Отдел закупки\ЛИЧНЫЕ\Семен\мус_внав.xlsx',
    sheet_name='Лист1')
    
def tg_norms():    
    #рассчитываем номера столбцов для каждого ТН
    if_string_lenth = pd.read_excel(r'C:\Остатки\наполняемость с 10012020.xlsx', sheet_name='ИФ', nrows=0, skiprows=2)
    
    lst = ['ИФ СЕРЬГИ' if 'ИФ СЕРЬГИ' in str(x) else x for x in if_string_lenth.columns.tolist()]
    
    if_end = len(lst)-lst[::-1].index('ИФ СЕРЬГИ')
    if_start = if_end-6
    if_columnrange = [1,2] + list(range(if_start,if_end))
    
    bk_string_lenth = pd.read_excel(r'C:\Остатки\наполняемость с 10012020.xlsx', sheet_name='БК', nrows=0, skiprows=2)
    
    lst = ['БК СЕРЬГИ' if 'БК СЕРЬГИ' in str(x) else x for x in bk_string_lenth.columns.tolist()]
    
    bk_end = len(lst)-lst[::-1].index('БК СЕРЬГИ')
    bk_start = bk_end-6
    bk_columnrange = [1,2] + list(range(bk_start,bk_end))
    
    # загружаем датафреймы
    begin = [1,2]
    
    if_df = pd.read_excel(
        r'C:\Остатки\наполняемость с 10012020.xlsx', sheet_name='ИФ', skiprows=3, usecols=if_columnrange, index_col=1,
        names=['блок', 'СКЛАД', 'ИФ КОЛЬЦА', 'ИФ КОЛЬЦА ОБРУЧ', 'ИФ ПЕЧАТКИ', 'ИФ ПОДВЕС ДЕКОР', 'ИФ ПОДВЕС КУЛЬТ', 'ИФ СЕРЬГИ'])
    
    bk_df = pd.read_excel(
        r'C:\Остатки\наполняемость с 10012020.xlsx', sheet_name='БК', skiprows=3, usecols=bk_columnrange, index_col=1,
        names=['блок', 'СКЛАД', 'БК КОЛЬЦА', 'БК КОЛЬЦА ОБРУЧ', 'БК ПЕЧАТКИ', 'БК ПОДВЕС ДЕКОР', 'БК ПОДВЕС КУЛЬТ', 'БК СЕРЬГИ'])
    
    # убираем закрытые маги
    if_df = if_df[~if_df.блок.str.contains('з', na=False)]
    del if_df['блок']
    
    bk_df = bk_df[~bk_df.блок.str.contains('з', na=False)]
    del bk_df['блок']
    
    #переворачиваем в нужный формат, соединяем и сохраняем
    if_tonav = priority_tonav(if_df)
    if_tonav.index.name = 'Код'
    
    bk_tonav = priority_tonav(bk_df)
    bk_tonav.index.name = 'Код'
    
    total = if_tonav.append(bk_tonav)
    
    total.to_excel(
        r'\\gold585.int\uk\Общее хранилище файлов\Коммерческий департамент\Отдел закупки\ЛИЧНЫЕ\Семен\приоритеты_внав.xlsx',
    sheet_name='Лист1')

def upload_norms():
    func_dict2 = {0: topnorms(), 1: wedring_norms(), 2: muslim_norms(), 3: normsforraiting()}
    return func_dict2[var_norms2.get()]
    
norms_b2 = Button(norms, text="Создать файл для загрузки в NAV", command=upload_norms)
norms_b2.pack(pady=10)  
    
n_book.add(stock, text='остатки')
n_book.add(distr, text='распределение товара')
n_book.add(orders, text='заказы')
n_book.add(process, text='обработка накладных')
n_book.add(norms, text='нормы')

window.mainloop()
