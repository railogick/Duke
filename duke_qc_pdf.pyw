# Python 3.7.2
import csv
import datetime
import os
import shutil

import pandas
import reportlab.lib.pagesizes
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas

now = datetime.datetime.now()


class VarFile:

    def __init__(self, filepath):
        # Initialize path information
        self._filepath = filepath
        self._win_path = '//Xmf-server/jobs'
        self._mac_path = '/Volumes/JOBS'

        # Initialize filename information
        self._fileName = os.path.basename(filepath)
        self._jobNumber = self._fileName[:8]
        self._jobName = self._fileName[:-4]
        self._jobExt = self._fileName[-4:]

        # Initialize pdf canvas and variables
        self.c = canvas.Canvas(f'{self._jobName} Checklist.pdf', pagesize=reportlab.lib.pagesizes.letter)
        self.c.translate(inch * .5, inch * .5)
        self.left_margin = 0
        self.right_margin = 7.5*inch

        # Initialize process variables
        self.record_count = 0
        self.proof_count = 0
        self.head_values = []
        self.record = []
        self.sample_dict = {}

        # import .csv file from BCC (as string to keep leading 0's in zip codes and other numerical fields)
        try:
            self.df = pandas.read_csv(self._filepath, engine='python', quotechar='"', sep=",", dtype=str)
        except pandas.errors.ParserError:
            self.df = pandas.read_csv(self._filepath, engine='python', quotechar='"', sep='\t', dtype=str)

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

        self.sample_dict = dict(zip(self.head_values, self.record))

        # Create new .csv file
        self.df.to_csv(self._jobName + '.csv', sep=',', quotechar='"',
                       quoting=csv.QUOTE_ALL, encoding='ISO-8859-1', index=False)

        # Place copy of file in the Data folder on the server.
        try:
            job_folder = [x for x in os.listdir(self._win_path) if x.startswith(self._jobNumber)]
            if job_folder:
                job_folder = f'{self._win_path}/{job_folder[0]}'
        except IOError:
            job_folder = [x for x in os.listdir(self._mac_path) if x.startswith(self._jobNumber)]
            if job_folder:
                job_folder = f'{self._mac_path}/{job_folder[0]}'
        else:
            if job_folder:
                data_folder = f'{job_folder}/Finals/Data'
                if os.path.exists(job_folder):
                    os.makedirs(data_folder, exist_ok=True)
                    shutil.copy2(self._jobName + '.csv', data_folder)

    def output_pdf(self):
        # set 1/2 inch margins
        self.header()
        self.body()
        self.footer()
        self.c.showPage()
        self.c.save()

    def header(self):
        self.c.line(0, 9.375*inch, 7.5*inch, 9.375*inch)
        self.c.setFont('Helvetica', 14)
        self.c.drawString(0, 10*inch, f'Job #: {self._jobNumber}')
        self.c.drawRightString(7.5*inch, 10*inch, f'Count: {str(len(self.df.index))}')
        self.c.setFont('Helvetica-Bold', 10)
        self.c.drawString(0, 9.5*inch, 'Database File:')
        self.c.setFont('Helvetica', 10)
        self.c.drawString(1*inch, 9.5*inch, f'{self._jobName}.csv')
        self.c.drawRightString(7.5*inch, 9.5*inch, f'Process Date: {now.strftime("%x")}')

        self.c.setFont('Helvetica-Bold', 18)
        self.c.drawCentredString(3.75*inch, 9.96875*inch, 'Variable Checklist')

    def body(self):
        col_one = self.left_margin
        y = 8.5*inch
        col_two = 3*inch  # Second Column
        size = 10
        self.c.setFont('Helvetica-Bold', size)
        self.c.drawString(col_one, y + 20, 'FIELD')
        self.c.line(col_one, y + 16, self.left_margin+1.375*inch, y + 16)
        self.c.drawString(col_two, y + 20, 'SAMPLE')
        self.c.line(col_two, y + 16, col_two+1.375*inch, y + 16)
        for key in self.sample_dict:
            self.c.setFont('Helvetica', size)
            self.c.drawString(self.left_margin, y, key)
            self.c.drawString(col_two, y, self.sample_dict[key])
            y = y - size*2
        if self.empty_columns:
            self.c.setFont('Helvetica-Bold', size)
            self.c.drawString(col_one, 1*inch, 'Empty Fields (Removed):')
            self.c.setFont('Helvetica', size)
            self.c.drawString(col_one, .75*inch, ', '.join(self.empty_columns))

    def footer(self):
        self.c.setFont('Helvetica', 11)
        self.c.drawString(self.left_margin, 0, 'Data Processed by: _________________________')
        self.c.drawRightString(self.right_margin, 0, 'QC by: _________________________')


def main():
    files = [p for p in os.listdir(os.getcwd()) if p.endswith(".csv") | p.endswith(".txt")]
    job = {}
    for idx, file in enumerate(files):
        try:
            job[idx] = VarFile(file)
        except pandas.errors.ParserError:
            print("Unable to process " + file)
            continue
        try:
            job[idx].process_file()
        except UnicodeEncodeError:
            with open("Process Error.txt", "w") as text_file:
                text_file.write('Encoding Error! Check for bad text in csv file\n')
                text_file.write('Example: (â€™) instead of standard apostrophe(\')')
            continue
        job[idx].output_pdf()
        del job[idx]


if __name__ == '__main__':
    main()