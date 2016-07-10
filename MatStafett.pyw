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
import openpyxl
import os
import tkinter
import random
import csv
from tkinter import filedialog
# from tkinter import messagebox


class Hmi:
    LANGUAGE = "eng"

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
        self.lang = {}

        # Get language pack
        self.get_lang(language=self.LANGUAGE)
        parent.title(self.lang["title"])

        # Main frames and labels
        self.l_title = tkinter.Label(parent, text=self.lang["title"])
        self.f_input = tkinter.Frame(parent, pady=10, padx=10)
        self.f_output = tkinter.Frame(parent, pady=10, padx=10)
        
        # TK Variables
        self.sv_filename = tkinter.StringVar()
        
        # Input widgets
        self.e_filename = tkinter.Entry(self.f_input, textvariable=self.sv_filename, width=30)
        self.b_select_file = tkinter.Button(self.f_input, text=self.lang["file_select"], command=self.select_file)
        self.b_run = tkinter.Button(self.f_input, text=self.lang["button_run"], command=self.generate_result,
                                    state=tkinter.DISABLED, height=3, width=10)

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
        self.scroll_x_output.grid(row=1, column=0, columnspan=2, sticky=tkinter.E + tkinter.W)
        self.scroll_y_output.grid(row=0, column=2, sticky=tkinter.N + tkinter.S)

    def select_file(self):
        """
        Method to open a file dialog and validate the file type.
        """
        file = filedialog.askopenfilename(title="Välj fil", initialdir=os.curdir,
                                          filetypes=self.list_supported_file_types)
        self.sv_filename.set(file)
        self.e_filename = self.sv_filename.get()
        file_ok = True
        if len(file) > 0:
            if file.endswith(".txt"):
                self.file_type = ".txt"
            elif file.endswith(".xlsx"):
                self.file_type = ".xlsx"
            else:
                file_ok = False
                self.log_output(self.lang["error_file_types"], "red")
        else:
            file_ok = False
            self.log_output(self.lang["error_no_file_selected"], "red")

        if file_ok:
            self.b_run.configure(state=tkinter.ACTIVE)
            self.file_path, self.file_name = os.path.split(file)
            self.log_output("{}: {}".format(self.lang["file_selected"], file))

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
        self.log_output(self.lang["file_reading_{}".format(self.file_type)])
        if self.file_type == ".txt":
            with open(file, "r") as f:
                for line in f:
                    self.list_participants.append(line)
        elif self.file_type == ".xlsx":
            wb = openpyxl.load_workbook(file, use_iterators=True)
            ws = wb.get_sheet_by_name(wb.sheetnames[0])
            max_rows = ws.max_row
            for row in ws.iter_rows("A1:A{}".format(max_rows)):
                for cell in row:
                    if cell.value is not None:
                        self.list_participants.append(cell.value)
        else:
            self.log_output(self.lang["error_file_types"], "red")

    def save_to_file(self):
        """
        Save the generated list to a file, Grouped and neatly ordered
        """
        self.log_output(self.lang["progress_gen_route"])
        starter = self.lang["starter"]+"\n"
        main = self.lang["main_course"]+"\n"
        desert = self.lang["desert"]+"\n"
        group = 0
        i = 1
        y = 2
        while group < self.num_groups:
            if i >= self.num_groups:
                i = 0
            if y >= self.num_groups:
                y = 0
            starter += "Värd:{}Gäster:\n{}{}\n".format(
                                                       self.groups_starter[group],
                                                       self.groups_main[group],
                                                       self.groups_desert[group])
            main += "Värd:{}Gäster:\n{}{}\n".format(
                                                    self.groups_main[i],
                                                    self.groups_starter[group],
                                                    self.groups_desert[y])
            desert += "Värd:{}Gäster:\n{}{}\n".format(
                                                      self.groups_desert[i],
                                                      self.groups_starter[group],
                                                      self.groups_main[y])
            group += 1
            i += 1
            y += 1
        self.log_output(self.lang["progress_done_saving"])
        # Todo fix xlsx...
        if self.file_type == ".txt" or self.file_type == ".xlsx":
            # Text file.
            # Create a new file since we dont want to mess with the source.
            filename = "new_" + self.file_name

            result = starter + main + desert
            # Open / create a new file and save the results.
            with open(os.path.join(self.file_path, filename), "w", encoding="utf8") as f:
                f.write(result)
            self.log_output("{} {}".format(self.lang["progress_saved_to"], filename))

    def validate_number_of_participants(self):
        """
        Makes sure the number of participants is a factor of 3.
        And at least 9
        Raises ValueError if not.
        """
        if len(self.list_participants) < 9:
            raise ValueError("Antal deltagare är mindre än 9")
        elif len(self.list_participants) % 3 != 0:
            raise ValueError("Antal deltagare är inte delbart på 3")
        else:
            self.num_groups = int(len(self.list_participants) / 3)
            self.log_output("{}: {}".format(self.lang["progress_found_participants"], len(self.list_participants)))

    def generate_random_index(self):
        self.log_output(self.lang["progress_gen_rand_list"])
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
        self.log_output(self.lang["progress_sort_unsort"])
        self.list_sorted_participants = []
        for index in self.list_rand_index:
            self.list_sorted_participants.append(self.list_participants[index-1])

        # create three equal sized lists containing all participants

        self.groups_starter = self.list_sorted_participants[0:self.num_groups]
        self.groups_main = self.list_sorted_participants[self.num_groups:self.num_groups*2]
        self.groups_desert = self.list_sorted_participants[self.num_groups*2:self.num_groups*3]

    def generate_result(self):
        """
        Start the process of generating the results.
        """
        self.read_file_contents()
        try:
            self.validate_number_of_participants()
        except ValueError as e:
            self.log_output("{}: {}".format(self.lang["error"], e))
            return
        self.generate_random_index()
        self.sort_participants()
        # Todo generate summary of who makes different parts of the meal
        self.save_to_file()

    def log_output(self, text, color="black"):
        """
        Method to print text to the output frame.
        :param text: The text to print
        :param color: Text color
        """
        self.t_output.configure(state=tkinter.NORMAL)
        text += "\n"
        self.t_output.insert(tkinter.END, text, color)
        self.t_output.configure(state=tkinter.DISABLED)

    def get_lang(self, language="eng", filename="lang.csv"):
        """
        Read a csv file and return the phrases matching the selected language
        :param language: the language to use
        :param filename: the csv file to open
        :return: a tuple with phrases in selected language
        """
        cur_lang = {}
        try:
            with open(filename, "r", encoding="utf8") as csv_file:
                reader = csv.DictReader(csv_file, delimiter=";")
                for row in reader:
                    cur_lang[row["phrase"]] = row[language]
        except KeyError:
            pass
        except FileNotFoundError:
            pass
            # Todo handle error
        self.lang = cur_lang

    def get_str(self, phrase):
        """
        Try to read a phrase from the language dictionary
        Throw an exception if not ok
        :param phrase: the wanted phrase
        :return: The text to print
        """
        try:
            string = self.lang[phrase]
            return string
        except KeyError:
            string = "Invalid phrase to print: {}".format(phrase)
            return string
        # Todo handle errors correctly

if __name__ == "__main__":
    root = tkinter.Tk()
    hmi = Hmi(root)
    hmi.LANGUAGE = "swe"
    hmi.draw_main()
    root.mainloop()
