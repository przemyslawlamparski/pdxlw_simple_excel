import numpy as np
import pandas as pd
import xlwings as xw
from datetime import date
import plotly.express as px
import plotly.figure_factory as ff

# https://www.kaggle.com/jackdaoud/marketing-data
df = pd.read_csv('C:\Marketing_campaign\marketing_campaign.csv', sep='\t')
print(df.head(5))
# load csv

app = xw.App(visible=False)
book = xw.Book()
sheet = book.sheets[0]
sheet.name = "Marketing_Campaing"
xw.sheets.active
# add new sheet

xw.sheets.add()
sheet1 = book.sheets[0]
xw.sheets.active
sheet1.name = "EDA"
# add second sheet

df = df.dropna()
# drop missing values

m = str(df['Marital_Status'].mode())
m = list(df.filter(regex='Mnt[A-Z]+'))
m = [i.replace('Mnt', '') for i in m]
m = [i.replace('Products', '') for i in m]
m = [i.replace('Prods', '') for i in m]
m

values = [df['MntWines'].sum(), df['MntFruits'].sum(), df['MntMeatProducts'].sum(), df['MntFishProducts'].sum()
          , df['MntSweetProducts'].sum(), df['MntGoldProds'].sum()]
fig = px.pie(df, values=values, names=m, color_discrete_sequence=px.colors.sequential.Teal)

fig.update_layout(legend=dict(
    yanchor="top",
    y=0.99,
    xanchor="left",
    x=0.85
))

fig.update_layout(plot_bgcolor='white',
                  font=dict(color="#909497"),
                  legend=dict(title='   Products sold    ',bordercolor="Black",borderwidth=1,bgcolor="#fff"),
                  title=dict(text='Breakdown by Producs',))

xw.sheets[0].pictures.add(fig, name='Prod', top=sheet1.range("A63").top, left=sheet1.range("A63").left, update=True)
# first pie plot

replace_values = {0:'No',1:'Yes'}
df['Complain'] = df['Complain'].replace(replace_values)
fig = px.pie(df, names='Complain', color_discrete_sequence=px.colors.sequential.Teal)
fig.update_traces(textinfo='label',textfont_size=15,textposition='auto')
fig.update_traces(hoverinfo='text+name')

fig.update_layout(legend=dict(
    yanchor="top",
    y=0.99,
    xanchor="left",
    x=0.922
))

fig.update_layout(plot_bgcolor='white',
                  font=dict(color="#909497"),
                  legend=dict(title='   Products sold    ',bordercolor="Black",borderwidth=1,bgcolor="#fff"),
                  title=dict(text='Breakdown by Complains',))

xw.sheets[0].pictures.add(fig, name='Compp', top=sheet1.range("M63").top, left=sheet1.range("M63").left, update=True)
# second pie plot

df['Dt_Customer'] =  pd.to_datetime(df['Dt_Customer'], infer_datetime_format=True)
df['Income'] = df['Income'].apply(np.int64)
df['Year_Birth'] = pd.to_datetime(df['Year_Birth'], format='%Y')
df['Age'] = df['Year_Birth'].apply(lambda x : (date.today().year - x.year))

d = np.datetime64('today')
df['Days_Enrolled'] =  d - df['Dt_Customer']
df['Days_Enrolled'] = (df['Days_Enrolled'] / np.timedelta64(1, 'D')).astype('int64')

cols = ['Age']
Q1 = df[cols].quantile(0.25)
Q3 = df[cols].quantile(0.75)
IQR = Q3 - Q1
df = df[~((df[cols] < (Q1 - 1.5 * IQR)) |(df[cols] > (Q3 + 1.5 * IQR))).any(axis=1)]

df.drop('Year_Birth', axis=1, inplace=True)
df['Total_Spending'] = df.filter(regex='Mnt[A-Z]+').sum(axis=1)

cols = ['Income', 'Total_Spending']
Q1 = df[cols].quantile(0.25)
Q3 = df[cols].quantile(0.75)
IQR = Q3 - Q1
df = df[~((df[cols] < (Q1 - 1.5 * IQR)) |(df[cols] > (Q3 + 1.5 * IQR))).any(axis=1)]

df["Marital_Status"] = df["Marital_Status"].replace(['Together'],'Married')
df["Marital_Status"] = df["Marital_Status"].replace(['Divorced','Widow','Alone','Absurd','YOLO'],'Single')
# clean up columns

hist_data = [df['Age']]
group_labels = ['Age']
fig= ff.create_distplot(hist_data, group_labels, histnorm='probability',
                        curve_type='normal', show_rug=False)

fig.update_layout(legend=dict(
    yanchor="top",
    y=0.99,
    xanchor="left",
    x=0.90
))

fig.update_layout(plot_bgcolor='white',
                  font=dict(color="#909497"),
                  title = dict(text = "Age distribution"),
                  legend=dict(bordercolor="Black",borderwidth=1,bgcolor="#fff"))

fig.data[0].showlegend = True

xw.sheets[0].pictures.add(fig, name='Aged', top=sheet1.range("Y37").top, left=sheet1.range("Y37").left, update=True)
# first histogram

df['Age'] = pd.cut(df['Age'],bins=11,labels=False,include_lowest=True)

hist_data = [df['Total_Spending']]
group_labels = ['Total Spending']
fig= ff.create_distplot(hist_data, group_labels, histnorm='probability', show_rug=False)

fig.update_layout(legend=dict(
    yanchor="top",
    y=0.99,
    xanchor="left",
    x=0.82
))

fig.update_layout(plot_bgcolor='white',
                  font=dict(color="#909497"),
                  title = dict(text = "Total Spending distribution"),
                  legend=dict(bordercolor="Black",borderwidth=1,bgcolor="#fff"))

xw.sheets[0].pictures.add(fig, name='Spendd', top=sheet1.range("M37").top, left=sheet1.range("M37").left, update=True)
# second histogram

hist_data = [df['Income'] / 1000 ]
group_labels = ['Income']
fig= ff.create_distplot(hist_data, group_labels, histnorm='probability',
                        curve_type='normal', show_rug=False)

fig.update_layout(legend=dict(
    yanchor="top",
    y=0.99,
    xanchor="left",
    x=0.88
))

fig.update_layout(plot_bgcolor='white',
                  font=dict(color="#909497"),
                  title = dict(text = "Income Distribution in thousands"),
                  legend=dict(bordercolor="Black",borderwidth=1,bgcolor="#fff"))



xw.sheets[0].pictures.add(fig, name='Incd', top=sheet1.range("A37").top, left=sheet1.range("A37").left, update=True)
# third histogram

fig = px.scatter(df, x="Total_Spending", y="Income", trendline="ols", trendline_color_override="green")
fig.update_traces(name = "OLS trendline")

fig.update_layout(
                  plot_bgcolor='white',
                  font=dict(color="#909497"),
                  title = dict(text = "Income against spending"),
                  xaxis = dict(title = "Spending", linecolor = "#909497"),
                  yaxis = dict(title = "Income", tickformat = ",", linecolor = "#909497"),
                  legend=dict(title='Marital Status',bordercolor="Black",borderwidth=1,bgcolor="#fff"))

fig.data[0].name = 'Spending'
fig.data[0].showlegend = True
fig.data[1].name = 'Income'
fig.data[1].showlegend = True

xw.sheets[0].pictures.add(fig, name='Inc', top=sheet1.range("A11").top, left=sheet1.range("A11").left, update=True)
# first scatter

fig = px.scatter(df, x="Total_Spending", y="Income",trendline="ols", trendline_color_override="green", color="Marital_Status")


fig.update_layout(
                  plot_bgcolor='white',
                  font=dict(color="#909497"),
                  title = dict(text = "Spendings against income by Marital Status"),
                  xaxis = dict(title = "Spending", linecolor = "#909497"),
                  yaxis = dict(title = "Income", tickformat = ",", linecolor = "#909497"),
                  legend=dict(title='Marital Status',bordercolor="Black",borderwidth=1,bgcolor="#fff"))

fig.data[0].name = 'Single'
fig.data[0].showlegend = True
fig.data[2].name = 'Married'
fig.data[2].showlegend = True
fig.data[1].name = 'Income'
fig.data[1].showlegend = True

xw.sheets[0].pictures.add(fig, name='Spend', top=sheet1.range("M11").top, left=sheet1.range("M11").left, update=True)
#second scatter

df['Children'] = df['Kidhome'] + df['Teenhome']
df.drop('Kidhome', axis=1, inplace=True)
df.drop('Teenhome', axis=1, inplace=True)

df.rename(columns={'MntWines':'Wines','MntFruits':'Fruits','MntMeatProducts':'MeatProducts',
          'MntFishProducts':'FishProducts','MntSweetProducts':'SweetProducts','MntGoldProds':'GoldProds'}, inplace=True)

#df.drop(df.filter(regex='Mnt[A-Z]+'), axis=1, inplace=True)
df.drop('Z_CostContact', axis = 1, inplace=True)
df.drop('Z_Revenue', axis=1, inplace=True)
df.drop('Dt_Customer', axis=1, inplace=True)

df = df[['ID', 'Age','Education', 'Children', 'Marital_Status', 'Days_Enrolled',  'Income', 'Total_Spending', 'Recency', 'Wines',
       'Fruits', 'MeatProducts', 'FishProducts', 'SweetProducts', 'GoldProds',
       'NumDealsPurchases', 'NumWebPurchases', 'NumCatalogPurchases',
       'NumStorePurchases', 'NumWebVisitsMonth', 'AcceptedCmp3',
       'AcceptedCmp4', 'AcceptedCmp5', 'AcceptedCmp1', 'AcceptedCmp2',
       'Complain', 'Response']]
#cleaning columns

count_freq = dict(df['Complain'].value_counts())
s = sum(count_freq.values())
l=[]
for k, v in count_freq.items():
    pct = str(round(v * 100.0 / s,2)) + '%'
    l.append(pct)

response  = str(round(df['Response'].mean() * 100,2)) + '%'
response

sheet1.range("A1").options(index=True).value = df['Age'].describe()
sheet1.range("D1").options(index=True).value = df['Days_Enrolled'].describe()
sheet1.range("G1").options(index=True).value = df['Income'].describe()
sheet1.range("J1").options(index=True).value = df['Total_Spending'].describe()

sheet1.range("M3").options(index=True).value = 'Campaing response rate: '
book.sheets[0].range('M3').api.Font.Bold = True
sheet1.range("N4").options(index=True).value = response

sheet1.range("P3").options(index=True).value = 'Complain rate: '
book.sheets[0].range('P3').api.Font.Bold = True
ount_freq = dict(df['Complain'].value_counts())
sheet1.range("R4").options(index=True).value = l
sheet1.range("R3").options(index=True).value = 'Yes'
sheet1.range("S3").options(index=True).value = 'No'
#more descriptive data

sheet.range("A1").options(index=False).value = df
book.save('C:\Marketing_campaign\Marketing_Campaing.xlsx')
excel_app = xw.apps.active
excel_app.quit()