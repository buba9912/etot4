import tkinter as tk
from tkinter import filedialog as fd

import ttk

from main import convert_to_t4_excel


def select_file():
    filetypes = (('excel files', '*.xlsx'),)
    input_filename = fd.askopenfilename(title='Open a file', initialdir='/', filetypes=filetypes)
    answer_type = 1
    xlsx_title = "Upload_file_t4_"
    convert_to_t4_excel(xlsx_title, answer_type, input_filename)


root = tk.Tk()
root.title('Excel to T4 readable File')
root.resizable(False, False)
root.geometry('480x440')
root.iconbitmap('icon.ico')
root.configure(background='#96bf0d')


# open button
open_button = ttk.Button(root, text='Open a File', command=select_file)
open_button.pack(expand=True, side=tk.LEFT)

# run the application
root.mainloop()
