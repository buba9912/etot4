import tkinter as tk
from tkinter import filedialog as fd

import ttk

from main import convert_to_t4_excel


def select_file():
    filetypes = (('excel files', '*.xlsx'),)
    input_filename = fd.askopenfilename(title='Open a file', initialdir='/', filetypes=filetypes)
    answer_type = 1
    xlsx_title = "Uploaded_file_t4_"
    convert_to_t4_excel(xlsx_title, answer_type, input_filename)


root = tk.Tk()
root.title('customer to t4 readable')
root.resizable(False, False)
root.geometry('600x350')


# open button
open_button = ttk.Button(root, text='Open a File', command=select_file)
open_button.pack(expand=True, side=tk.LEFT)

# run the application
root.mainloop()
