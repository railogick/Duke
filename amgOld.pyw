# Python 3.7.2
from datetime import datetime

from gooey import Gooey, GooeyParser
from pandas import read_excel, read_csv, DataFrame, errors

now = datetime.now()

# TODO: Retrieve Job Number
# TODO: Trim leading/trailing whitespaces from contract numbers. DONE
# TODO: Fix Merge Summary
# TODO: Remove Envelope column from processed data list. DONE
# TODO: Fix output file names.


class AnthemMerge:
    """
    Two arguments regulate the purpose of the script:
    1. DATA - generates counts of the two versions of envelopes and checks for mismatched data and separates the list
    by 'Envelope' for presorting.
    2. MERGE - Appends branding grid data to the presorted lists and adds proofs to the top the list. The proofs
    consist of the first 2 records of each Contract Number.
    """
    def __init__(self, brandgrid, maillist):

        # initialize DataFrames
        self._dfMerged = DataFrame()  # Joined Branding Grid and Mailing List.
        self._dfProofs = DataFrame()  # Contract Number samples.
        self._dfFinal = DataFrame()  # Final DataFrame for output
        self._dfMailGrid = DataFrame()  # Compressed Branding Grid (Contract Number, Envelope only)
        self._dfMailMerge = DataFrame()  # Mailing list to be grouped by Envelope for presort.

        # Import Branding Grid and Mailing List
        try:
            self._dfGrid = read_excel(brandgrid, dtype=str)
        except errors.ParserError:
            print("Unable to process Branding Grid")
        try:
            self._dfList = read_csv(maillist, dtype=str)
        except errors.ParserError:
            print("Unable to process Mailing List")

        # ---- Branding Grid cleanup ---- #
        self._dfGrid.columns = self._dfGrid.columns.str.title().str.strip()
        # Assign Anthem to any empty 'Envelope' fields
        self._dfGrid['Envelope'] = self._dfGrid['Envelope'].fillna('Anthem')
        # Fill in remaining na values.
        self._dfGrid = self._dfGrid.fillna('')
        self._dfGrid = self._dfGrid.apply(lambda x: x.str.strip())

        # Merge Contract Number columns and drop original columns.
        print('Generating Contract Numbers...')
        self._dfGrid['Contract Number'] = self._dfGrid.apply(self.col_merge, axis=1)
        self._dfGrid.drop(['Cms Contract', '2018 Pbp', 'Sourcegroupnumber', 'Sourcesubgrpnbr'],
                          axis=1,
                          inplace=True)

    # --- Data Adjustment Functions ---
    @staticmethod
    def col_merge(row):
        """
        Combine Contract Elements separated by a '-'
        :param row:
        :return str:
        """
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

    def list_check(self, row):
        """
        Check for 4 tier List Contract Number in grid_df.
        If not found, check for 3 tier, otherwise return 2 tier List Contract Number.
        :param row:
        :return str:
        """
        if row['List Contract Number'] in row[self._dfGrid['Contract Number']]:
            return row['List Contract Number']
        if "-".join(row['List Contract Number'].split('-')[:3]) in row[self._dfGrid['Contract Number']]:
            return "-".join(row['List Contract Number'].split('-')[:3])
        else:
            return "-".join(row['List Contract Number'].split('-')[:2])

    def grid_check(self, row):
        """
        Method to assign partial Contract Number matches
        Check for Contract Number in _dfList. If not found, reduce to 2 tier Contract Number
        :param row:
        :return str:
        """
        if row['Contract Number'] in row[self._dfList['Contract Number']]:
            return row['Contract Number']
        else:
            return "-".join(row['Contract Number'].split('-')[:2])

    # --- Create joined DataFrames ---
    def merge(self):
        print('Merging files...')
        self._dfList.drop(['Envelope'], axis=1, inplace=True)  # Drop Envelope Column from presorted list before merge.
        self._dfGrid['Contract Number'] = self._dfGrid.apply(self.grid_check, axis=1)
        self._dfMerged = self._dfList.join(self._dfGrid.drop_duplicates(['Contract Number'])
                                           .set_index('Contract Number'),
                                           on='Contract Number')

        self._dfFinal = self._dfMerged  # set to _dfFinal for when proofs aren't being made.

    def get_proofs(self):
        """
        Generate a new data working_df consisting of the first 2 records
        of each unique contract number
        """

        print('Getting unique Contract Numbers for proofs..')
        contracts = self._dfMerged['Contract Number'].unique()
        for x in range(len(contracts)):
            self._dfProofs = self._dfProofs.append(self._dfMerged.loc[self._dfMerged['Contract Number'] ==
                                                                      contracts[x]].head(2))
        self._dfProofs['Proofs'] = 'Proof'
        print('Adding proofs to mailing list...')
        self._dfFinal = self._dfProofs.append(self._dfMerged, ignore_index=True, sort=False)

    # --- Output Methods ---
    def mail_list_env(self):
        # Create working_df for Envelope Merge
        self._dfMailGrid = self._dfGrid[['Contract Number', 'Envelope']]

        # Cleanup Mailing List
        self._dfList.columns = self._dfList.columns.str.title()
        print('Verifying Contract Numbers...')
        self._dfList['Contract Number'] = self._dfList.apply(self.list_check, axis=1)
        # self._dfMailGrid['Contract Number'] = self._dfMailGrid.apply(self.grid_check, axis=1)
        self._dfMailMerge = self._dfList.join(self._dfMailGrid.drop_duplicates(['Contract Number'])
                                              .set_index('Contract Number'),
                                              on='Contract Number')

        # Output new .csv files based on Envelope type.
        sorted_frames = dict(tuple(self._dfMailMerge.groupby(['Envelope'])))
        frame_list = [sorted_frames[x] for x in sorted_frames]
        for idx, frame in enumerate(frame_list):
            frame_list[idx].to_csv(frame_list[idx].iloc[0]['Envelope'] + f' Envelope List_{now:%m%d%y}.csv',
                                   index=False,
                                   header=True)

        # Output breakdown of merged list based on Envelope type and Contract Number
        # Fill na values with #N/A to indicate missing values from join.
        df_group = self._dfMailMerge.fillna('#N/A').groupby(['Envelope', 'Contract Number'])
        df_group['Envelope'].agg(len).to_csv(f'Merge Summary_{now:%m%d%y}.csv', header=True)

    def create_csv(self):
        """ Output to csv file with Latin(ISO-8859-1) encoding for compatibility with Variable Data Software """
        name = f'Anthem Merged_{now:%m%d%y}'
        print('Outputting list to .csv...')
        self._dfFinal.to_csv(name + '.csv', index=False, encoding='ISO-8859-1')


@Gooey(program_name='Anthem Merge Program')
def main():
    parser = GooeyParser(description='Combine the branding grid with a processed mailing list for variable data merge')
    parser.add_argument('Branding_Grid',
                        help="Select the Branding Grid File (.xlsx)",
                        widget="FileChooser")
    parser.add_argument('Mailing_List',
                        help="Select the Mailing List File (.csv)",
                        widget="FileChooser")
    purpose_group = parser.add_argument_group(title='Purpose:')
    radio_group = purpose_group.add_mutually_exclusive_group()
    radio_group.add_argument("-Data", help="Generate Data for Mail Processing. Use with Customer Data.",
                             action="store_true")
    radio_group.add_argument("-Merge", help="Combine processed mail list with Branding Grid. Use with Presorted Data.",
                             action="store_true")
    args = parser.parse_args()

    anthemjob = AnthemMerge(args.Branding_Grid, args.Mailing_List)
    if args.Data:
        # Process files for data processing. For use with customer original data.
        anthemjob.mail_list_env()
    if args.Merge:
        # Process files for mail merge. For use with presorted data.
        anthemjob.merge()
        anthemjob.get_proofs()
        anthemjob.create_csv()
    del anthemjob
    print('Job Complete!')


if __name__ == '__main__':
    main()