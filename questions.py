import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import FormatStrFormatter

sales = pd.read_excel("data.xlsx")  # чтение таблицы

# 1)Вычисление общей выручки за июль 2021, приход денежных средств которых не просрочен

a = sales[sales['status'] == 'Июль 2021'].index[0]  # выбор промежутка интересующих нас данных из таблицы
b = sales[sales['status'] == 'Август 2021'].index[0]

month = sales[a + 1:b]
summ = month[month['status'] != "ПРОСРОЧЕНО"]['sum'].sum()  # подсчет выручки за выбранный период

print('Общая выручка за июль 2021: ', round(summ, 2))

# 2)Изменение выручки компании за выбранный период по месяцам, построение графика

x = list(
    filter(lambda i: not i.isupper(), sales['status'].values[:]))  # составление списка месяцов рассматриваемого периода

idxA = sales[sales['status'].isin(x)].index.to_list()  # разметка таблицы по месяцам
idxB = idxA[1:] + [len(sales)]

y = []
for i in range(len(x)):  # подсчет выручки за каждый отдельный месяц
    month = sales[idxA[i] + 1:idxB[i]]
    y.append(round(float(month['sum'].sum()), 2))

# построение графика
fig = plt.figure(figsize=(15, 7))
ax = fig.add_subplot()
ax.set_title('Изменение выручки')
ax.yaxis.set_major_formatter(FormatStrFormatter('% d'))
ax.set_ylabel('Выручка')
ax.set_xlabel('Период')
ax.plot(x, y)
ax.grid()
plt.show()

# 3) кто из менеджеров привлек больше денежных средств в сентябре 2021

a = idxA[x.index('Сентябрь 2021')]  # выбор промежутка интересующих нас данных из таблицы
b = idxB[x.index('Сентябрь 2021')]

month = sales[a + 1:b]

grouped = month.groupby(['sale'])  # группирование выборки по менеджерам
managers = grouped.agg({'sum': 'sum'})  # подсчет суммы для каждой группы
manager = managers.loc[managers.idxmax()].index[0]  # получение фамилии менеджера с самой большей суммой
print(f"Менеджер {manager} привлек больше всего денежных средств в сентябре 2021")

# 4)Какой тип сделок (новая/текущая) был преобладающим в октябре 2021?
a = idxA[x.index('Октябрь 2021')]  # выбор промежутка интересующих нас данных из таблицы
b = idxB[x.index('Октябрь 2021')]

month = sales[a + 1:b]

grouped = month.groupby(['new/current']).count()  # группирование и подсчет количесства данных в каждой группе
print("В октябре 2021 преобладающим типом сделок был: ", grouped.loc[grouped.idxmax()].index[0])

# 5)Сколько оригиналов договора по майским сделкам было получено в июне 2021?
a = idxA[x.index('Май 2021')]  # выбор промежутка интересующих нас данных из таблицы
b = idxB[x.index('Май 2021')]

month = sales[a + 2:b]
# поиск записей в рассматриваемом периоде, подходящих под условия
second = month[(month['receiving_date'] >= pd.Timestamp(year=2021, month=6, day=1)) & (
        month['receiving_date'] <= pd.Timestamp(year=2021, month=6, day=30))].count()

print(second.loc[second.index[0]], " оригиналов договора по майским сделкам было получено в июне 2021")
