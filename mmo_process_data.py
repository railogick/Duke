import re
from datetime import datetime
from os import path, listdir, mkdir, rename
from xml.etree.ElementTree import parse

from gooey import Gooey, GooeyParser
from numpy import NaN
from pandas import ExcelWriter, DataFrame, read_excel, concat, set_option

now = datetime.now()

# -------------- Class Definitions -----------------------


class XmlImport:
    def __init__(self, xml_dir):
        """
        Parses xml data received from Medical Mutual of Ohio online information request site.
        :rtype: object
        """
        # Initialization of variables
        self.xml_dir = xml_dir
        self.total_orders = []
        self.order = []
        self.df = DataFrame()

        # Definition of XML Data conversion to proper header names.
        self.header_dict = {'ShipToName': 'Full Name', 'ShipToAddress1': 'Address', 'ShipToCity': 'City',
                            'ShipToState': 'State', 'ShipToZip': 'Zip', 'ShipToAddress3': 'Zone',
                            'ProductCode': 'Product Code', 'ProductName': 'Product Desc', 'PromiseDate': 'Drop Date',
                            'OrderType': 'Order Type', 'OrderDate': 'Order Date', 'UserEmail': 'Email',
                            'ShipToAddress4': 'Phone', 'BillToRegion': 'Bill To Region', 'PlanYear': 'PlanYear',
                            'PlanType': 'PlanType', 'MemberType': 'MemberType',
                            'WebtrendscampaignIDcode': 'WebtrendscampaignIDcode'}

    def parse_xml(self):
        """
        Open XML File and convert orders into a Dataframe.
        """
        for filename in listdir(self.xml_dir):
            if not filename.endswith('.xml'):
                continue
            fullname = path.join(self.xml_dir, filename)
            tree = parse(fullname)
            root = tree.getroot()

            # Get information from each order and add to total_orders
            for each in root.iter('Order'):
                for x in range(len(self.header_dict)):
                    self.order.append(each.find(f'.//{list(self.header_dict.keys())[x]}').text)

                # Add completed order to total_orders and reset order list for next file.
                self.total_orders.append(self.order)
                self.order = []

        # Create Data frame from completed total_orders
        self.df = DataFrame(self.total_orders, columns=self.header_dict.values())
        self.df.insert(0, 'BRC_ID', value='')

        # Check to see if files were processed, and if so, put them in a dated folder.
        if not self.df.empty:
            folder = f'{self.xml_dir}/XML {now:%m%d%y}'
            if not path.exists(folder):
                mkdir(folder)
            for f in listdir(self.xml_dir):
                if f.endswith('.xml'):
                    rename(f'{self.xml_dir}/{f}', f'{folder}/{f}')

    def xml_to_xlsx(self):
        """
        Convert xml file into a spreadsheet.
        """
        if not self.df.empty:
            xml_filename = f'MMO_XML_ORDER {now:%m-%d-%Y}.xlsx'
            writer = ExcelWriter(xml_filename, engine='xlsxwriter')
            self.df.to_excel(writer, sheet_name='Sheet1', startrow=1, header=False, index=False)
            ws = writer.sheets['Sheet1']
            ws.set_column('A:R', 18)
            rows = int(len(self.df.index))
            table_headers = []
            for idx, val in enumerate(self.df.columns.values):
                table_headers.append(dict(header=val))
            ws.add_table(0, 0, rows, 17,
                         {'columns': table_headers})
            writer.save()


class ProcessFile:
    def __init__(self, frame):
        set_option('precision', 0)
        self.df = frame
        self.updates = {}

        # Find and remove empty data
        self.df.replace(' ', NaN, inplace=True)
        self.df.dropna(subset=['Full Name', 'Address'], inplace=True)

        # Process data
        self.update()
        self.remove_dupes()
        self.output_single_year()
        # self.separate_by_year()

    def update(self):
        """
        Dictionary of Updates
        Format - Column to Change:{Text to Change: New Text}
        :return: N/A
        """
        self.updates = {'Product Code': {'MMO ENROLLKIT': 'PEK',
                                         'MMO MGNFR': 'MAG',
                                         'MMO UMG': 'UMG',
                                         NaN: 'UMG',
                                         '': 'UMG'},
                        'Product Desc': {'Optional Supplemental Benefit (OSB) Fulfillment Kit': '(OSB) Fulfillment Kit',
                                         NaN: 'Understanding Medicare Guide',
                                         '': 'Understanding Medicare Guide'},
                        'Order Type': {'SalesCallCenter': 'Call Center',
                                       'End User': 'WEB',
                                       'CustomerCare': 'Customer Service',
                                       NaN: 'BRE',
                                       '': 'BRE'},
                        'Bill To Region': {'Region 1': '1',
                                           'Region 1S': 'R1S',
                                           'Region 2': '2',
                                           NaN: '2',
                                           '': '2'},
                        # 'PlanYear': {NaN: str(datetime.now().year), '': str(datetime.now().year)},
                        'PlanYear': {NaN: '2020', '': '2020', '2019': '2020'}
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
        self.df.drop_duplicates(subset=['Full Name', 'Address'], inplace=True)

        # Remove test records
        self.df = self.df[~self.df['Full Name'].str.lower().str.contains('test')]

        # Start the index at 1 and sort by 'Product Code'
        self.df.reset_index(drop=True, inplace=True)
        self.df.index += 1
        self.df = self.df.sort_values(by='Product Code')

    def output_single_year(self):
        writer = ExcelWriter(str(self.df.iloc[0]['PlanYear']) + ' list.xlsx')
        self.df.to_excel(writer, index=False, header=True)
        ws = writer.sheets['Sheet1']
        # rows = int(len(df_list[idx].index))
        self.df.columns = map(str.upper, self.df.columns)
        ws.write_row(0, 0, self.df.columns.values)
        writer.save()

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
            ws = writer.sheets['Sheet1']
            # rows = int(len(df_list[idx].index))
            df_list[idx].columns = map(str.upper, df_list[idx].columns)
            ws.write_row(0, 0, df_list[idx].columns.values)
            writer.save()


# -------------- Functions ---------------------------

def output_contact_dnc(contact_df, dnc_df):
    filename = f'MMO CONTACT & DO NOT MAIL {now:%m-%d-%Y}.xlsx'
    writer = ExcelWriter(filename, engine='xlsxwriter')
    contact_df.to_excel(writer, sheet_name='Sheet1', startrow=2, header=False, index=False)
    ws = writer.sheets['Sheet1']
    ws.set_column('A:H', 20)
    ws.write('A1', f'{now:%m/%d/%Y}')
    contact_df.columns = map(str.upper, contact_df.columns)
    ws.write_row('A2', list(contact_df.columns.values))
    if not dnc_df.empty:
        dnc_row = len(contact_df.index) + 4
        ws.write(dnc_row, 0, 'DO NOT MAIL')
        dnc_df.to_excel(writer, sheet_name='Sheet1', startrow=dnc_row + 1, header=False, index=False)
    contact_df.columns = map(str.title, contact_df.columns)
    writer.save()


@Gooey
def main():
    parser = GooeyParser()
    parser.add_argument('-Data_File',
                        help="Select data entry file",
                        widget="FileChooser")
    args = parser.parse_args()

    # Process XML file
    xml_dir = '//Xmf-server/duke/Inter Office Mail/MMO XML Orders/'
    mmo_xml = XmlImport(xml_dir)
    mmo_xml.parse_xml()
    mmo_xml.xml_to_xlsx()

    # Create main working DataFrame
    mmo_df = mmo_xml.df

    # Process Data Entry file if present
    if args.Data_File:
        data_entry_df = read_excel(args.Data_File, dtype=object)
        df_contact = data_entry_df.loc[data_entry_df['STATUS'] == 'CONTACT'].drop(columns=['STATUS']).reset_index(drop=True)
        df_dnc = data_entry_df.loc[data_entry_df['STATUS'] == 'DNC'].drop(columns=['STATUS']).reset_index(drop=True)
        df_data = data_entry_df.loc[data_entry_df['STATUS'] == 'LIST'].drop(columns=['STATUS']).reset_index(drop=True)

        # output Contact List and Do Not Contact list to excel
        output_contact_dnc(df_contact, df_dnc)

        # Add Contact List and Data List to main DataFrame.
        mmo_df = concat([mmo_df, df_contact, df_data],
                        ignore_index=True, sort=False).drop(columns=['Check Box']).fillna('')

    job = ProcessFile(mmo_df)


if __name__ == '__main__':
    main()
