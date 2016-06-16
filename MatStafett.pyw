"""
Project: Matstafett
Author: Patrik Viljehav
Version: 1.0
Description:    Select an Excel (.xlsx) or notepad (.txt) document.
                The script will calculate a setup where all participants
                will:
                    * Get a certain part of the dinner to prepare.
                    * Meet new people at every part of the dinner
                    * Save the result in the same file type as was submitted to the script (this can be changed)
"""

# Import needed modules
# import openpyxl
import os
import tkinter
from tkinter import filedialog
# from tkinter import messagebox


class Hmi:
    def __init__(self, parent):
        """
        Initialize the class
        :param parent: The master tk window
        """
        # variables
        self.file_type = ""
        self.file_name = ""
        self.list_supported_file_types = [("Excel", "xlsx"), ("Text", "txt")]

        self.l_title = tkinter.Label(parent, text="Matstafett")
        self.f_input = tkinter.Frame(parent, pady=10, padx=10)
        self.f_output = tkinter.Frame(parent, pady=10, padx=10)
        # TK Variables
        self.sv_filename = tkinter.StringVar()

        # Input widgets
        self.e_filename = tkinter.Entry(self.f_input, textvariable=self.sv_filename, width=30)
        self.b_select_file = tkinter.Button(self.f_input, text="Välj fil", command=self.select_file)
        self.b_run = tkinter.Button(self.f_input, text="KÖR!", command=self.generate_result, state=tkinter.DISABLED,
                                    height=3, width=10)

        # Output widgets
        self.t_output = tkinter.Text(self.f_output, height=10, width=70, state=tkinter.DISABLED)

    def draw_main(self):
        """
        Draw the main window
        """
        # Main title
        self.l_title.grid(row=0, column=0)

        # input frame
        self.f_input.grid(row=1, column=0)
        self.e_filename.grid(row=0, column=0)
        self.b_select_file.grid(row=0, column=1)
        self.b_run.grid(row=1, column=0, columnspan=2)

        # output frame
        self.f_output.grid(row=2, column=0)
        self.t_output.grid(row=0, column=0, columnspan=2)

    def select_file(self):
        """
        Method to open a file dialog and validate the file type.
        """
        file = filedialog.askopenfilename(title="Välj fil", initialdir=os.curdir,
                                          filetypes=self.list_supported_file_types)
        file_ok = True
        if len(file) > 0:
            if file.endswith(".txt"):
                self.file_type = ".txt"
            elif file.endswith(".xlsx"):
                self.file_type = ".xlsx"
            else:
                file_ok = False
                self.log_output("Kan bara öppna .txt eller .xlsx (Excel)-filer.")
        else:
            file_ok = False
            self.log_output("Ingen fil vald :(")

        if file_ok:
            self.b_run.configure(state=tkinter.ACTIVE)
            self.file_name = file
            self.log_output("Vald Fil: {}".format(file))
        else:
            self.b_run.configure(state=tkinter.DISABLED)
            self.file_name = ""
            self.file_type = ""

    def generate_result(self):
        pass

    def log_output(self, text):
        """
        Method to print text to the output frame.
        :param text: The text to print
        """
        self.t_output.configure(state=tkinter.NORMAL)
        text += "\n"
        self.t_output.insert(tkinter.END, text)
        self.t_output.configure(state=tkinter.DISABLED)

if __name__ == "__main__":
    root = tkinter.Tk()
    root.title("Matstafett")
    hmi = Hmi(root)
    hmi.draw_main()
    root.mainloop()
