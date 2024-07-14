import pandas as pd

df = pd.read_excel('practice-parsing/output/microwave_articles.xlsx', index_col=None)

df['Date'] = pd.to_datetime(df['Date'])

df = df[df['Date'].dt.year >= 2020]

df.to_excel('practice-parsing/output/microwave_articles.xlsx', engine='xlsxwriter', index=False)