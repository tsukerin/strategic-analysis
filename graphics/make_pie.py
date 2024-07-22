import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import pygal

df = pd.read_csv('found-keywords/output/domains_keyword_counts.csv')

pie = pygal.Pie(print_values=True, inner_radius=.4)
pie.title = 'Процентное распределение областей применения'

for i, j in df.iterrows():
    pie.add(j['Obj'], j['Count'] / sum(df['Count']) * 100, formatter=lambda x: str(round(x)) + '%')

pie.render_to_png('graphics/domains_pie.png')