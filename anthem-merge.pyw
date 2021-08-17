from datetime import datetime
from os import path

import pandas as pd
from gooey import Gooey, GooeyParser
from xlsxwriter import Workbook

now = datetime.now()


def initialize_branding_grid(branding_grid_file):
    grid_df = pd.DataFrame()
    try:
        grid_df = pd.read_excel(branding_grid_file, dtype=str)
    except pd.errors.ParserError:
        print("Unable to process Branding Grid")

    grid_df.columns = grid_df.columns.str.title().str.strip()
    grid_df['Envelope'] = grid_df['Envelope'].fillna('Anthem')  # Fill in envelope data that is not otherwise indicated.
    grid_df = grid_df.fillna('')
    grid_df = grid_df.apply(lambda x: x.str.strip())

    # Merge Contract Number columns
    grid_df['Contract Number'] = grid_df.apply(branding_grid_col_merge, axis=1)
    grid_df.drop_duplicates(subset=['Contract Number'], inplace=True)
    grid_df.set_index('Contract Number', inplace=True)
    grid_df.drop(['Cms Contract', '2018 Pbp', 'Sourcegroupnumber', 'Sourcesubgrpnbr'], axis=1, inplace=True)
    return grid_df


def branding_grid_col_merge(row):
    contract_name = []
    if row['Cms Contract']:
        contract_name.append(row['Cms Contract'])
    if row['2018 Pbp']:
        contract_name.append(row['2018 Pbp'])
    if row['Sourcegroupnumber']:
        contract_name.append(row['Sourcegroupnumber'])
    if row['Sourcesubgrpnbr']:
        contract_name.append(row['Sourcesubgrpnbr'])
    return '-'.join(contract_name)


def initialize_mailing_list(mailing_list_file, grid_df):
    list_df = pd.DataFrame()
    try:
        list_df = pd.read_csv(mailing_list_file, dtype=str)
    except pd.errors.ParserError:
        print("Unable to process Mailing List")

    # Create Tier 3 and Tier 2 Contract Numbers for matching with the Branding Grid.
    list_df['LCN3'] = list_df['List Contract Number'].str.rsplit('-', 1, expand=True)[0]
    list_df['LCN2'] = list_df['List Contract Number'].str.rsplit('-', 2, expand=True)[0]

    # Find matching Contract Numbers in the Branding Grid.
    list_df.loc[list_df['List Contract Number'].isin(grid_df.index),
                'Contract Number'] = list_df['List Contract Number']
    list_df.loc[list_df['LCN3'].isin(grid_df.index), 'Contract Number'] = list_df['LCN3']

    # Fill in remainder with Tier 2 contract numbers.
    list_df.loc[~list_df[['List Contract Number', 'LCN3']].isin(grid_df.index).any(axis=1),
                'Contract Number'] = list_df['LCN2']

    # Drop temporary columns
    list_df.drop(['LCN3', 'LCN2'], axis=1, inplace=True)
    return list_df


def output_envelope_csv(merge_frame):
    anthem_df = merge_frame.loc[merge_frame['Envelope'] == 'Anthem']
    amerigroup_df = merge_frame.loc[merge_frame['Envelope'] == 'Amerigroup']
    anthem_df.to_csv(f'Anthem Envelope List {now:%m%d%y}.csv', header=True, index=False)
    amerigroup_df.to_csv(f'Amerigroup Envelope List {now:%m%d%y}.csv', header=True, index=False)


def list_summary(merge_frame):
    cname = {'City': 'Count'}
    df_group = merge_frame.rename(columns=cname).fillna('#N/A').groupby(['Envelope', 'Contract Number'],
                                                                        as_index=False).count()
    envelope_list = ['#N/A', 'Amerigroup', 'Anthem']
    dfs = []
    for idx, frame in enumerate(envelope_list):
        dfs.append(df_group[df_group['Envelope'] == frame]
                   [['Envelope', 'Contract Number', 'Count']].reset_index(drop=True))

    with Workbook(f'Anthem Merge Summary_{now:%m%d%y}.xlsx') as wb:
        ws = wb.add_worksheet()
        fmt_header = wb.add_format({'font_size': 14, 'bold': 1, 'align': 'center'})
        fmt_bold = wb.add_format({'bold': 1})
        fmt_right = wb.add_format({'align': 'right'})
        col_iter = iter(range(len(dfs)))
        for idx, frame in enumerate(dfs):
            col = int(idx*2+next(col_iter))
            ws.set_column(col, col, 26)
            ws.set_column(col+1, col+1, 9, fmt_right)
            ws.merge_range(0, col, 0, col+1, frame['Envelope'][0], fmt_header)
            ws.write(1, col, 'Total', fmt_bold)
            ws.write(1, col+1, frame['Count'].agg(sum), fmt_bold)
            ws.write_row(2, col, frame.columns[1:3].values)
            ws.write_column(3, col, frame['Contract Number'])
            ws.write_column(3, col+1, frame['Count'])


def generate_proofs(working_df):
    contracts = working_df['Contract Number'].unique()
    proof_df = pd.DataFrame()
    for x in range(len(contracts)):
        proof_df = proof_df.append(working_df.loc[working_df['Contract Number'] == contracts[x]].head(2))
    proof_df['Proofs'] = 'Proof'
    proof_df = proof_df.append(working_df, ignore_index=True, sort=False)
    return proof_df


@Gooey(program_name='Anthem Merge Program')
def main():
    parser = GooeyParser(description='Combine the branding grid with a processed mailing list for variable data merge')
    parser.add_argument('Branding_Grid',
                        help="Select the Branding Grid File (.xlsx)",
                        widget="FileChooser")
    parser.add_argument('Mailing_List',
                        help="Select the Mailing List File (.csv)",
                        widget="FileChooser")
    parser.add_argument('Purpose',
                        choices=['Data', 'Merge'],
                        default='Data',
                        widget='Dropdown',
                        help='Select purpose of merge:')

    args = parser.parse_args()
    grid_df = initialize_branding_grid(args.Branding_Grid)
    mail_df = initialize_mailing_list(args.Mailing_List, grid_df)

    if 'Data' in args.Purpose:
        env_df = grid_df['Envelope']
        merged_df = mail_df.join(env_df, on='Contract Number')
        output_envelope_csv(merged_df)
        list_summary(merged_df)

    if 'Merge' in args.Purpose:
        mail_df.drop(['Envelope'], axis=1, inplace=True)
        merged_df = mail_df.join(grid_df, on='Contract Number')
        final_df = generate_proofs(merged_df)
        filename = path.basename(args.Mailing_List).rsplit(' ', 1)[0]
        final_df.to_csv(f'{filename} Merged Variable.csv', index=False, encoding='ISO-8859-1')


if __name__ == '__main__':
    main()