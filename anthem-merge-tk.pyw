# Python 3.7.2
import datetime
from os import path
from tkinter import *
from tkinter.filedialog import askopenfilename
from tkinter.ttk import *

from amgOld import AnthemMerge

now = datetime.datetime.now()

# TODO: Change import based on file type.
# TODO: Cleanup the interface
# TODO: Retrieve Job Number
# TODO: Comment code.


class Window(Frame):
    def __init__(self, master=None, **kw):
        Frame.__init__(self, master)
        super().__init__(master, **kw)
        self.master = master
        self.pack(fill=BOTH, expand=1)

        self.brandgrid = ''
        self.maillist = ''

        lbl_brand = Label(self, text='Branding Grid: ')
        lbl_list = Label(self, text='Mailing List: ')
        lbl_job = Label(self, text='Job Number: ')

        lbl_brand.grid(column=0, row=0)
        lbl_list.grid(column=0, row=1)
        lbl_job.grid(column=0, row=2)

        self.ent_brand = Entry(self)
        self.ent_list = Entry(self)
        self.ent_job = Entry(self)

        self.ent_brand.grid(column=1, row=0)
        self.ent_list.grid(column=1, row=1)
        self.ent_job.grid(column=1, row=2)

        self.btn_brand = Button(self, text="Select", command=self.get_grid)
        self.btn_list = Button(self, text="Select", command=self.get_list)
        self.btn_proc = Button(self, text="Run Job", command=self.process)

        self.btn_brand.grid(column=2, row=0)
        self.btn_list.grid(column=2, row=1)
        self.btn_proc.grid(column=0, row=4)

    def get_grid(self):
        self.brandgrid = askopenfilename(filetypes=[('Excel Files', '*.xlsx')])

        # Clear Entry Field if already populated.
        if self.brandgrid:
            self.ent_brand.delete(0, END)

        self.ent_brand.insert(0, path.basename(self.brandgrid))

    def get_list(self):
        self.maillist = askopenfilename(filetypes=[('CSV Files', '.csv')])

        # Clear Entry Field if already populated.
        if self.maillist:
            self.ent_list.delete(0, END)

        self.ent_list.insert(0, path.basename(self.maillist))

    def process(self):
        anthemjob = AnthemMerge(self.brandgrid, self.maillist)
        anthemjob.merge()
        anthemjob.get_proofs()
        anthemjob.create_csv()
        del anthemjob


def main():
    root_frame = Tk()
    root_frame.title('Anthem Merge')
    root_frame.geometry('400x300')
    app = Window(root_frame)
    root_frame.mainloop()


if __name__ == '__main__':
    main()