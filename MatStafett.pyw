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
import random
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
        self.file_name = None
        self.file_path = None
        self.list_supported_file_types = [("Excel", "xlsx"), ("Text", "txt")]
        self.list_participants = []
        self.list_sorted_participants = []
        self.groups_starter = []
        self.groups_main = []
        self.groups_desert = []
        self.list_rand_index = []
        self.num_groups = None
        
        # Main frames and labels
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
        # Add colors and scrollbars to the output widget
        log_colors = ["black", "green", "blue", "red"]
        for color in log_colors:
            self.t_output.tag_config(color, foreground=color)
        self.scroll_x_output = tkinter.Scrollbar(self.f_output, command=self.t_output.xview, orient=tkinter.HORIZONTAL)
        self.scroll_y_output = tkinter.Scrollbar(self.f_output, command=self.t_output.yview)
        self.t_output.configure(yscrollcommand=self.scroll_y_output.set, xscrollcommand=self.scroll_x_output.set)

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
            self.file_path, self.file_name = os.path.split(file)
            self.log_output("Vald Fil: {}".format(file))
        else:
            self.b_run.configure(state=tkinter.DISABLED)
            self.file_name = None
            self.file_path = None
            self.file_type = ""

    def read_file_contents(self):
        """
        Read the participant list in the selected file.
        """
        self.list_participants = []
        file = os.path.join(self.file_path, self.file_name)
        self.log_output("Läser in {}-filen för att få lite koll på vilka som vill vara med.".format(self.file_type))
        if self.file_type == ".txt":
            with open(file, "r") as f:
                for line in f:
                    self.list_participants.append(line)
        elif self.file_type == ".xlsx":
            self.log_output("Excel file to be implemented...")
        else:
            self.log_output("Filtyp måste vara txt eller xlsx")

    def save_to_file(self):
        """
        Save the generated list to a file, Grouped and neatly ordered
        """
        starter = ""
        group = 0
        while group < self.num_groups:
            starter += "Grupp {} \n{}{}{}\n".format(group + 1,
                                                    self.groups_starter[group][0],
                                                    self.groups_starter[group][1],
                                                    self.groups_starter[group][2])
            group += 1

        if self.file_type == ".txt":
            # Text file.
            # Create a new file since we dont want to mess with the source.
            filename = "new_" + self.file_name
            # Prepare a variable to write to the file.
            # participants = ""
            # for participant in self.list_participants:
            #     participants += participant

            # Open / create a new file and save the results.
            with open(os.path.join(self.file_path, filename), "w") as f:
                f.write(starter)

    def validate_number_of_participants(self):
        """
        Makes sure the number of participants is a factor of 3.
        And at least 9
        Raises ValueError if not.
        """
        self.log_output(str(len(self.list_participants)))
        if len(self.list_participants) < 9:
            raise ValueError("Antal deltagare är mindre än 9")
        elif len(self.list_participants) % 3 != 0:
            raise ValueError("Antal deltagare är inte delbart på 3")
        else:
            self.num_groups = len(self.list_participants) / 3

    def generate_random_index(self):
        i = 1
        self.list_rand_index = []
        while i <= len(self.list_participants):
            self.list_rand_index.append(i)
            i += 1
        random.shuffle(self.list_rand_index)

    def sort_participants(self):
        """
        Generate a sorted... sort of unsorted... list of members based on random numbers.
        """
        self.list_sorted_participants = []
        for index in self.list_rand_index:
            self.list_sorted_participants.append(self.list_participants[index-1])

        group = 1
        i = 0
        self.groups_starter = []
        while group <= self.num_groups:
            self.groups_starter.append([self.list_sorted_participants[i],
                                        self.list_sorted_participants[i+1],
                                        self.list_sorted_participants[i+2]])
            group += 1
            i += 3

    def generate_result(self):
        """
        Start the process of generating the results.
        """
        self.read_file_contents()
        try:
            self.validate_number_of_participants()
        except ValueError as e:
            self.log_output("Error: {}".format(e))
            return
        self.generate_random_index()
        self.sort_participants()

        # todo save unsordet list as stage 1
        # todo keep first 3. shift all else 3 times.

        self.save_to_file()

    def log_output(self, text, color="black"):
        """
        Method to print text to the output frame.
        :param text: The text to print
        :param color: Text color
        """
        # todo add scrollbar
        self.t_output.configure(state=tkinter.NORMAL)
        text += "\n"
        self.t_output.insert(tkinter.END, text, color)
        self.t_output.configure(state=tkinter.DISABLED)

if __name__ == "__main__":
    root = tkinter.Tk()
    root.title("Matstafett")
    hmi = Hmi(root)
    hmi.draw_main()
    root.mainloop()
