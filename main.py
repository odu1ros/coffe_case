import pandas as pd

df = pd.read_csv('data.csv', delim_whitespace=True, decimal=",")
df = df[df.Delivery != 1]


df['Number Of Transactions'] = 1
df['Date'] = df['TransactionID']+' '+df['Date']
df.Date = pd.to_datetime(df.Date)

df = df.sort_values(by="Date")

df = df.set_index('Date').groupby([pd.Grouper(freq="M"), "Promo"]).sum()

df['Profitability, %'] = (df['Cost']-df['Paid'])/(df['Cost'])*100
df['Bad review, %'] = df['BadReview']/df['Number Of Transactions']*100

df = df.drop(columns=['Count', 'Cost', 'Delivery', 'Paid', 'Weight', 'AutoIssue'])
daf = pd.DataFrame(df, columns= ['Number Of Transactions', 'Bad review, %', 'Profitability, %'])

daf.to_excel("output.xlsx", sheet_name='results')
