import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo

# create the root window
from main import new_xlsx, refactor_file, main, convert_to_t4_excel

root = tk.Tk()
root.title('customer to t4 readable')
root.resizable(False, False)
root.geometry('600x350')



def select_file():
    filetypes = (
        ('excel files', '*.xlsx'),
        # ('All files', '*.*')
    )

    inputFilename = fd.askopenfilename(
        title='Open a file',
        initialdir='/',
        filetypes=filetypes)

    xlsx_title = "gui"
    answer_type = 1

    convert_to_t4_excel(xlsx_title, answer_type, inputFilename)



# open button
open_button = ttk.Button(
    root,
    text='Open a File',
    command=select_file
)

open_button.pack(expand=True)

# run the application
root.mainloop()
