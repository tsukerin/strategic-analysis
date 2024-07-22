import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import pygal

df = pd.read_csv('found-keywords/output/attrs_keyword_counts.csv')

bar = pygal.HorizontalBar()
bar.title = 'Наиболее популярные атрибуты в регионе (Европа)'

for i, j in df.iterrows():
    bar.add(j['Obj'], j['Count'])

bar.render_to_png('graphics/attrs_bars.png')