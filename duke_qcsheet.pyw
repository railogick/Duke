  
# Python 3.7.2
from csv import QUOTE_ALL
from datetime import datetime
from os import getcwd, path, listdir, makedirs
from shutil import copy2
from typing import List, Dict

from pandas import read_csv, errors
from xlsxwriter import Workbook


class VarFile:
    """
    PURPOSE: Import mailing list for cleanup and generate a job information sheet in Excel.
    1. Remove completely empty Columns/Rows
    2. Remove columns with no data.
    3. Rename Header duplicates.
    4. Remove leading/trailing whitespace in data
    5. Get sample record with most populated fields.
    6. Output data sheet to Excel with job information and sample record
    7. Copy data file to xmf-server job folder if present (For Mac and PC Operating Systems).
    """

    def __init__(self, filepath):
        # Set Job information
        self._filepath = filepath
        self._fileName = path.basename(filepath)
        self._jobNumber = self._fileName[:8]
        self._jobName = self._fileName[:-4]
        self._jobExt = self._fileName[-4:]
        self._win_path = '//Xmf-server/jobs'
        self._mac_path = '/Volumes/JOBS'
        self.record_count = 0
        self.proof_count = 0
        self.head_values = []
        self.record = []

        try:
            self.df = read_csv(self._filepath,
                               engine='python',
                               quotechar='"',
                               sep=",",
                               dtype=str)  # dtype str to keep leading 0's
        except errors.ParserError:
            self.df = read_csv(self._filepath,
                               engine='python',
                               quotechar='"',
                               sep='\t',
                               dtype=str)  # dtype str to keep leading 0's

        # Create list of empty columns that will be dropped
        self.empty_columns = self.df.columns[self.df.isna().all()].tolist()

    def process_file(self):
        """
        Cleanup data frame and generate sample data for excel sheet
        """
        self.df.dropna(how='all', inplace=True)
        self.df.dropna(axis=1, how='all', inplace=True)
        self.df = self.df.apply(lambda x: x.str.strip())
        self.head_values = self.df.columns.values

        # Get record with all fields populated for sample
        self.record = self.df.dropna(how='any').head(1).values.tolist()

        # If full record doesn't exist, get next full record minus pkg/con columns.
        if not self.record:
            empty = self.df.columns.values[self.df.isin(['###']).any() |
                                           self.df.isin(['***']).any()]
            sublist = [x for x in self.head_values if x not in empty]

            # Loop to find longest record based on least amount of na values.
            x = 1
            while not self.record:
                self.record = self.df.dropna(subset=sublist,
                                             thresh=len(self.head_values) - x).fillna('').head(1).values.tolist()
                x += 1

        # Cleanup record list for display in Excel (Limit chars to 40 to fit cells)
        self.record = self.record[0]
        for x in range(len(self.record)):
            self.record[x] = self.record[x][0:41]

        # Create new .csv file
        self.df.to_csv(self._jobName + '.csv',
                       sep=',',
                       quotechar='"',
                       quoting=QUOTE_ALL,
                       encoding='ISO-8859-1',
                       index=False)

        # Place copy of file in the Data folder on the server.
        try:
            job_folder = [x for x in listdir(self._win_path) if x.startswith(self._jobNumber)]
            if job_folder:
                job_folder = f'{self._win_path}/{job_folder[0]}'
        except IOError:
            job_folder = [x for x in listdir(self._mac_path) if x.startswith(self._jobNumber)]
            if job_folder:
                job_folder = f'{self._mac_path}/{job_folder[0]}'
        else:
            if job_folder:
                data_folder = f'{job_folder}/Finals/Data'
                if path.exists(job_folder):
                    makedirs(data_folder, exist_ok=True)
                    copy2(self._jobName + '.csv', data_folder)

    def output_files(self):

        # Create Excel sheet
        with Workbook(f'{self._jobName} Checklist.xlsx') as wb:
            ws = wb.add_worksheet()

            # Formatting
            now = datetime.now()
            fmt_bold = wb.add_format({'bold': 1})
            fmt_title = wb.add_format({'bottom': True,
                                       'bold': 1,
                                       'font_size': 13})
            fmt_head_border = wb.add_format({'top': True})

            # Print specifications
            ws.fit_to_pages(1, 0)  # Fit to 1x1 pages.
            ws.set_page_view()

            # Worksheet Formatting
            ws.hide_gridlines(2)
            ws.set_column('A:E', 17)
            ws.set_default_row(20)

            # Header
            ws.set_margins(top=1.125)
            ws.set_header(f'&L&16Job #: {self._jobNumber}&11\n\n'
                          f'&\"Calibri,Bold\"Database File: &\"Calibri,Regular\"{self._jobName}.csv'
                          f'&C&\"Calibri,Bold\"&18Variable Checklist&R&16Count: '
                          f'{str(len(self.df.index))}&11\n\nProcess Date: {now.strftime("%x")}')
            ws.merge_range('A1:E1', '', fmt_head_border)

            # Footer
            ws.set_footer('&L&\"Calibri,Bold\"Data Processed by: _________________________'
                          '&R&\"Calibri,Bold\"QC by: _________________________')

            # Write data to worksheet
            ws.write('A3', 'FIELD', fmt_title)
            ws.write('C3', 'SAMPLE', fmt_title)
            ws.write_column('A4', self.head_values)
            ws.write_column('C4', self.record)

            # List Removed Empty Columns if any.
            if self.empty_columns:
                start1 = 'A' + str(9 + len(self.head_values))
                start2 = 'A' + str(10 + len(self.head_values))
                ws.write(start1, 'Empty Fields (Removed):', fmt_bold)
                ws.write(start2, ', '.join(self.empty_columns))


def main():
    files: List[str] = [p for p in listdir(getcwd())
                        if p.endswith(".csv") | p.endswith(".txt")]

    job: Dict[int, VarFile] = {}
    for idx, file in enumerate(files):
        try:
            job[idx] = VarFile(file)
        except errors.ParserError:
            print("Unable to process " + file)
            continue
        try:
            job[idx].process_file()
        except UnicodeEncodeError:
            with open("Process Error.txt", "w") as text_file:
                text_file.write('Encoding Error! Check for bad text in csv file\n')
                text_file.write('Example: (â€™) instead of standard apostrophe(\')')
            continue
        job[idx].output_files()
        del job[idx]


if __name__ == '__main__':
    main()