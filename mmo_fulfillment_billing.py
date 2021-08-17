# Python 3.7.2
""" Script to Automate the End of Month Summary and Billing Tasks"""
# -----------------------
# -- Imports/Functions --
# -----------------------
# Test information
from os import getcwd, listdir
from re import search

from numpy import select, where, ceil
from pandas import ExcelWriter, read_excel, pivot_table, DataFrame, concat

# Package Postage Cost
PEK: float = 8.30
MAG: float = 1.84
OSB: float = 2.05
OSB_PRS: float = 2.68
UMG: float = 1.84

# Reply Envelope Postage
BRE: float = 0.64
# Data Entry Charge
DATA: float = 0.20
# Fulfillment Charge
CHARGE: float = 0.52


def job_summary(job_frame):
    """ Pivot Job to Count Product Codes """
    job_frame = pivot_table(job_frame,
                            index=['PRODUCT CODE'],
                            values=['ADDRESS'],
                            aggfunc=len,
                            margins=True)
    job_frame.reset_index(inplace=True)
    job_frame.columns = ['Product Code', 'Count']
    job_frame['Product Code'].replace(to_replace=['All'],
                                      value='Grand Total', inplace=True)
    return job_frame


def get_col(i):
    """ Take index and alternate column assignment
    :param i:
    :return x:
    """
    if i % 2 == 0:
        x = 0
    else:
        x = 4
    return x


def get_row(current_idx):
    """ Starting row based on index.
    :param current_idx:
    :return current_row:
    """
    start_row = int(ceil((current_idx-1)/2))  # Formula to set 2 frames per row.
    max_height = 8  # max height of DataFrame summary
    spacer = 3      # initial start row & space between DataFrames
    current_row = start_row * max_height + spacer
    return current_row

# -----------------------------------
# -- Retrieve Files for Processing --
# -----------------------------------


file = []
for p in listdir(getcwd()):
    if search(r'W\d{1,2}[A-E]', p):
        file.append(p)
job_num = file[0][0:8]

# ------------------------------------
# -- Workbook & Worksheet Formatting --
# ------------------------------------

writer = ExcelWriter(f'{job_num} EOM.xlsx', engine='xlsxwriter')
wb = writer.book

# Formatting for Excel worksheets
fmt_breakdown = wb.add_format({'align': 'center', 'italic': True})
fmt_center = wb.add_format({'align': 'center'})
fmt_currency = wb.add_format({'num_format': '$#,##0.00'})
fmt_header = wb.add_format({'bold': True, 'border': 1})
fmt_sum_head = wb.add_format({'align': 'center', 'border': 1,
                              'num_format': '$#,##0.00'})
fmt_sum_count = wb.add_format({'align': 'center', 'border': 1})
fmt_ttl = wb.add_format({'bold': True, 'align': 'right'})
fmt_sub_ttl = wb.add_format({'bold': True, 'top': 1})
# Orange hex color: 'ffc000'

# ------------------
# -- Process Data --
# ------------------

df = DataFrame()

# Loop through files to generate summary and append to billing.

for idx, item in enumerate(file):
    temp = read_excel(item)
    col = get_col(idx)
    row = get_row(idx)
    df_sum = job_summary(temp)
    df_sum.to_excel(writer,
                    sheet_name='Summary',
                    header=False,
                    startcol=col,
                    startrow=row,
                    index=False)
    heads = list(df_sum.columns.values)
    sub_ttl = row + len(df_sum.index) - 1
    ws_summary = writer.sheets['Summary']
    ws_summary.merge_range(row - 3, col, row - 3, col + 1,
                           item[9:13] + ' Breakdown',
                           fmt_breakdown)
    ws_summary.write(row - 1, col + 1,
                     'Total',
                     fmt_ttl)
    ws_summary.conditional_format(sub_ttl, col, sub_ttl, col + 1,
                                  {'type': 'no_blanks',
                                   'format': fmt_sub_ttl})
    df = concat([df, temp], ignore_index=True, sort=False)

# Determine OSB Prospect mailings
df.loc[(df['PRODUCT CODE'] == 'MMO OSB') &
       (df['WEBTRENDSCAMPAIGNIDCODE'] == 'Prospect'),
       'PRODUCT CODE'] = 'MMO OSB PROSPECT'

# Assign values to product codes
CONDITIONS = [
    (df['PRODUCT CODE'] == 'PEK'),
    (df['PRODUCT CODE'] == 'MMO MAG'),
    (df['PRODUCT CODE'] == 'MMO OSB'),
    (df['PRODUCT CODE'] == 'MMO OSB PROSPECT'),
    (df['PRODUCT CODE'] == 'UMG')]
CHOICES = [PEK, MAG, OSB, OSB_PRS, UMG]

# Create calculated columns
df.insert(0, 'Postage In', where(df['ORDER TYPE'] == 'BRE', BRE, 0))
df.insert(1, 'Postage Out', select(CONDITIONS, CHOICES))
df.insert(2, 'Data Entry', where(df['ORDER TYPE'] == 'BRE', DATA, 0))
df.insert(3, 'Fulfillment Charge', CHARGE)
df.insert(4, 'TTL', df['Postage In'] + df['Postage Out'] +
          df['Data Entry'] + df['Fulfillment Charge'])
df.insert(5, 'Count', 1)

# Add BRE Junk Envelopes Received
junk_mail = int(input("Total Junk Mail received: "))
junk_post_in = junk_mail * BRE
df = df.append({'Postage In': junk_post_in, 'TTL': junk_post_in, 'Count': junk_mail},
               ignore_index=True, sort=False)

# Get column totals for placement at top of worksheet
sum_row = df[['Postage In', 'Postage Out', 'Data Entry', 'Fulfillment Charge',
              'TTL', 'Count']].sum()
df_sum = DataFrame(data=sum_row).T
df_sum = df_sum.reindex(columns=df.columns)
sums = df_sum.iloc[[0][0:6]].values.flatten().tolist()

# --------------------------------
# -- Output/Format Billing Data --
# --------------------------------

ws_summary.set_column('A:A', 20)
ws_summary.set_column('E:E', 20)

df.to_excel(writer, sheet_name='Billing', startrow=4,
            header=False, index=False)
ws_billing = writer.sheets['Billing']
ws_billing.set_column('A:E', 18, fmt_currency)

heads = list(df.columns.values)
for idx, item in enumerate(heads):
    ws_billing.write(3, idx, item, fmt_center)

ws_billing.conditional_format('A4:F4', {'type': 'no_blanks',
                                        'format': fmt_header})

ws_billing.write_row(2, 0, sums[0:5], fmt_sum_head)
ws_billing.write(2, 5, sums[5], fmt_sum_count)
ws_billing.write('E1', 'Total Junk Mail:')
ws_billing.write('F1', junk_mail)
ws_billing.write_formula('F2', '=F3-F1')

wb.close()
© 2021 GitHub, Inc.
Terms
Privacy
Security
Status
Docs
Contact GitHub
Pricing
API
Training
Blog
About
