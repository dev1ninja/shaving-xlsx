import pandas as pd

file_path = '../test.xlsx'

allow_job_titles = ['Managing Director', 'Co-Head', 'COO', 'Director', 'Chief Executive', 'Co-Founder', 'CFO', 'CEO', 'CIO', 'Chief Financial Officer', 'Co-CEO', 'Chief Investor Relations Officer', 'CCO', 'Chief Talent Officer', 'Chief Operating Partner', 'COO', 'Chief Information Officer', 'Co-Investment', 'Chief Legal Officer', 'Chief Compliance Officer']

columns = ['FIRM_ID', 'CONTACT_ID', 'FUND MANAGER', 'FIRM TYPE', 'TITLE', 'NAME', 'LinkedIn', 'ALTERNATIVE NAME', 'JOB TITLE', 'ASSET CLASS', 'EMAIL', 'TEL', 'CITY', 'STATE', 'COUNTRY/TERRITORY', 'ZIP CODE']

column = ('FIRM_ID', 'CONTACT_ID', 'FUND MANAGER', 'FIRM TYPE', 'TITLE', 'NAME', 'LinkedIn', 'ALTERNATIVE NAME', 'JOB TITLE', 'ASSET CLASS', 'EMAIL', 'TEL', 'CITY', 'STATE', 'COUNTRY/TERRITORY', 'ZIP CODE')

data = {
    'FIRM_ID': [],
    'CONTACT_ID': [],
    'FUND MANAGER': [],
    'FIRM TYPE': [],
    'TITLE': [],
    'NAME': [],
    'LinkedIn': [],
    'ALTERNATIVE NAME': [],
    'JOB TITLE': [],
    'ASSET CLASS': [],
    'EMAIL': [],
    'TEL': [],
    'CITY': [],
    'STATE': [],
    'COUNTRY/TERRITORY': [],
    'ZIP CODE': []
}

df = pd.read_excel(r'../test.xlsx', 'Contacts_Export', engine='openpyxl', dtype=object, header=None)
print('file read finished!')

count_row = df.shape[0]
count_col = df.shape[1]

print('Row: ', count_row, 'Col: ', count_col)

for row_index in range(count_row):
    for index in allow_job_titles:
        if str(index) in str(df[8][row_index]):
            # print(index, df[8][row_index])
            for col_index in range(count_col):
                data[columns[col_index]].append(df[col_index][row_index])
# df1 = pd.DataFrame(data[0:10000])
# df2 = pd.DataFrame(data[10000:20000])
# df3 = pd.DataFrame(data[20000:30000])
# df4 = pd.DataFrame(data[30000:40000])
# df5 = pd.DataFrame(data[40000:50000])
# df6 = pd.DataFrame(data[50000:60000])
# df7 = pd.DataFrame(data[60000:70000])

# writer = pd.ExcelWriter('beta.xlsx', engine='xlsxwriter')

# df1.to_excel(writer, sheet_name='Sheet1')
# df2.to_excel(writer, sheet_name='Sheet2')
# df3.to_excel(writer, sheet_name='Sheet3')
# df4.to_excel(writer, sheet_name='Sheet4')
# df5.to_excel(writer, sheet_name='Sheet5')
# df6.to_excel(writer, sheet_name='Sheet6')
# df7.to_excel(writer, sheet_name='Sheet7')

# writer.save()

df1 = pd.DataFrame(data=data)
df1.to_csv('./beta.csv')
# df1.to_excel('./beta.xlsx')
print('file write finished!')
# print(data)
# df.to_excel('./beta.xlsx')

# print(df[8])