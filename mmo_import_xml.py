# Python 3.7.2
from os import path, listdir, mkdir, rename
from xml.etree.ElementTree import parse
from datetime import datetime
from pandas import ExcelWriter, DataFrame

now = datetime.now()


class XmlImport:
    def __init__(self, xml_dir):
        """
        Parses xml data received from Medical Mutual of Ohio online information request site.
        :rtype: object
        """
        self.xml_dir = xml_dir
        self.total_orders = []
        self.order = []
        self.df = DataFrame()
        self.header_dict = {'ShipToName': 'Full Name', 'ShipToAddress1': 'Address', 'ShipToCity': 'City',
                            'ShipToState': 'State', 'ShipToZip': 'Zip', 'ShipToAddress3': 'Zone',
                            'ProductCode': 'Product Code', 'ProductName': 'Product Desc', 'PromiseDate': 'Drop Date',
                            'OrderType': 'Order Type', 'OrderDate': 'Order Date', 'UserEmail': 'Email',
                            'ShipToAddress4': 'Phone', 'BillToRegion': 'Bill To Region', 'PlanYear': 'PlanYear',
                            'PlanType': 'PlanType', 'MemberType': 'MemberType',
                            'WebtrendscampaignIDcode': 'WebtrendscampaignIDcode'}

    def parse_xml(self):
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

                # Add completed order to total_orders and reset order list
                self.total_orders.append(self.order)
                self.order = []

        # Create Data frame from completed total_orders
        self.df = DataFrame(self.total_orders, columns=self.header_dict.values())
        self.df.insert(0, 'BRC_ID', value='')

        if not self.df.empty:
            folder = f'{self.xml_dir}/XML {now:%m%d%y}'
            if not path.exists(folder):
                mkdir(folder)
            for f in listdir(self.xml_dir):
                if f.endswith('.xml'):
                    rename(f'{self.xml_dir}/{f}', f'{folder}/{f}')

    def xml_to_xlsx(self):
        if not self.df.empty:
            xml_filename = f'MMO_XML_ORDER {now:%m-%d-%Y}.xlsx'
            writer = ExcelWriter(xml_filename, engine='xlsxwriter')
            self.df.to_excel(writer, sheet_name='Sheet1', startrow=1, header=False, index=False)
            wb = writer.book
            ws = writer.sheets['Sheet1']
            ws.set_column('A:R', 18)
            rows = int(len(self.df.index))
            ws.add_table(0, 0, rows, 17,
                         {'columns': [{'header': 'BRC_ID'}, {'header': 'Full Name'}, {'header': 'Address'},
                                      {'header': 'City'}, {'header': 'State'}, {'header': 'Zip'},
                                      {'header': 'Zone'}, {'header': 'Product Code'},
                                      {'header': 'Product Desc'}, {'header': 'Drop Date'},
                                      {'header': 'Order Type'}, {'header': 'Order Date'},
                                      {'header': 'Email'}, {'header': 'Phone'},
                                      {'header': 'Bill To Region'}, {'header': 'PlanYear'},
                                      {'header': 'PlanType'}, {'header': 'MemberType'},
                                      {'header': 'WebtrendscampaignIDcode'},
                                      ]})
            writer.save()


def main():
    xml_dir = '//Xmf-server/duke/Inter Office Mail/MMO XML Orders/'
    xml_df = XmlImport(xml_dir)
    xml_df.parse_xml()
    xml_df.xml_to_xlsx()


if __name__ == '__main__':
    main()