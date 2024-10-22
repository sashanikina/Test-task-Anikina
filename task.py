import pandas as pd
import numpy as np

sales = pd.read_excel("data.xlsx")  # чтение таблицы
x = list(
    filter(lambda x: not x.isupper(), sales['status'].values[:]))  # составление списка месяцов рассматриваемого периода

idxA = sales[sales['status'].isin(x)].index.to_list()  # разметка таблицы по месяцам
idxB = idxA[1:] + [len(sales)]
managers = np.unique(np.concatenate([sales['sale'].values[idxA[0] + 2:idxB[0]], sales['sale'].values[idxA[1] + 1:idxB[
    1]]]))  # выбор значений в рассматриваемом периоде
nulls = [0] * len(managers)
dic = dict(zip(managers, nulls))  # составление словаря менеджера и его остатка

# подсчет остатков каждого менеджера по периодам
for i in range(idxA[0] + 2, idxB[0]):
    if sales['new/current'].values[i] == 'новая' and sales['status'].values[i] == 'ОПЛАЧЕНО' and \
            sales['receiving_date'].values[i]:
        if sales['receiving_date'].values[i] >= pd.Timestamp(year=2021, month=7, day=1):
            dic[sales['sale'].values[i]] += float(0.07 * sales['sum'].values[i])
    elif sales['new/current'].values[i] == 'текущая' and sales['status'].values[i] != 'ПРОСРОЧЕНО' and \
            sales['receiving_date'].values[i]:
        if sales['receiving_date'].values[i] >= pd.Timestamp(year=2021, month=7, day=1):
            if sales['sum'].values[i] >= 10000:
                dic[sales['sale'].values[i]] += float(0.05 * sales['sum'].values[i])
            else:
                dic[sales['sale'].values[i]] += float(0.03 * sales['sum'].values[i])

for i in range(idxA[1] + 2, idxB[1]):
    if sales['document'].values[i] != 'НЕТ':
        if sales['new/current'].values[i] == 'новая' and sales['status'].values[i] == 'ОПЛАЧЕНО' and \
                sales['receiving_date'].values[i]:
            if sales['receiving_date'].values[i] >= pd.Timestamp(year=2021, month=7, day=1):
                dic[sales['sale'].values[i]] += float(0.07 * sales['sum'].values[i])
        elif sales['new/current'].values[i] == 'текущая' and sales['status'].values[i] != 'ПРОСРОЧЕНО' and \
                sales['receiving_date'].values[i]:
            if sales['receiving_date'].values[i] >= pd.Timestamp(year=2021, month=7, day=1):
                if sales['sum'].values[i] >= 10000:
                    dic[sales['sale'].values[i]] += float(0.05 * sales['sum'].values[i])
                else:
                    dic[sales['sale'].values[i]] += float(0.03 * sales['sum'].values[i])

for k, v in dic.items():
    print(f"У менеджера {k} на 01.07.2021 наблюдается остаток {round(v, 2)}")
