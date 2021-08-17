# Python 3.7.2
import re
from datetime import datetime

from numpy import NaN
from pandas import set_option, read_excel, ExcelWriter


class ProcessFile:
    def __init__(self, frame):
        set_option('precision', 0)
        self.df = frame
        self.updates = {}

        # Find and remove empty data
        self.df.replace(' ', NaN, inplace=True)
        self.df.dropna(subset=['Full Name'], inplace=True)
        self.df.dropna(subset=['Address'], inplace=True)

        # Process data
        self.update()
        self.remove_dupes()
        self.separate_by_year()

    def update(self):
        """
        Dictionary of Updates
        Format - Column to Change:{Text to Change: New Text}
        :return: N/A
        """
        self.updates = {'Product Code': {'MMO ENROLLKIT': 'PEK',
                                         'MMO MGNFR': 'MAG',
                                         'MMO UMG': 'UMG',
                                         NaN: 'UMG'},
                        'Product Desc': {'Optional Supplemental Benefit (OSB) Fulfillment Kit': '(OSB) Fulfillment Kit',
                                         NaN: 'Understanding Medicare Guide'},
                        'Order Type': {'SalesCallCenter': 'Call Center',
                                       'End User': 'WEB',
                                       'CustomerCare': 'Customer Service',
                                       NaN: 'BRE'},
                        'Bill To Region': {'Region 1': '1',
                                           'Region 2': '2',
                                           NaN: '2'},
                        'PlanYear': {NaN: datetime.now().year}
                        }

        self.df.replace(self.updates, inplace=True)

    def remove_dupes(self):
        """
        Clean data by converting to title case and trimming whitespaces before checking for duplicates.
        Reset index and sort by 'Product Code'
        :return: N/A
        """
        fix_cols = ['Full Name', 'Address', 'City']
        self.df[fix_cols] = self.df[fix_cols].applymap(lambda x: x.title())
        self.df[fix_cols] = self.df[fix_cols].applymap(lambda x: re.sub(' +', ' ', x))
        self.df.drop_duplicates(['Full Name', 'Address'], inplace=True)

        # Remove test records
        self.df = self.df[~self.df['Full Name'].str.lower().str.contains('test')]

        # Start the index at 1 and sort by 'Product Code'
        self.df.reset_index(drop=True, inplace=True)
        self.df.index += 1
        self.df.sort_values(by='Product Code', inplace=True)

    def separate_by_year(self):
        """
        Split DataFrame into 2 output files, grouped by 'PlanYear'
        :return:
        """
        df_dict = dict(tuple(self.df.groupby(['PlanYear'])))
        df_list = [df_dict[x] for x in df_dict]
        for idx, frame in enumerate(df_list):
            writer = ExcelWriter(str(df_list[idx].iloc[0]['PlanYear']) + ' list.xlsx')
            df_list[idx].to_excel(writer, index=False, header=True)
            wb = writer.book
            ws = writer.sheets['Sheet1']
            rows = int(len(df_list[idx].index))
            df_list[idx].columns = map(str.upper, df_list[idx].columns)
            ws.write_row(0, 0, df_list[idx].columns.values)
            writer.save()


def main():
    now = datetime.now()
    filename = f'//Xmf-server/duke/Inter Office Mail/Medical Mutual Spreadsheets/MMO Fulfillment/_IN PROCESS/MMO_XML_ORDER {now:%m-%d-%Y}.xlsx'
    xml_df = read_excel(filename, dtype=str)
    job = ProcessFile(xml_df)


if __name__ == '__main__':
    main()