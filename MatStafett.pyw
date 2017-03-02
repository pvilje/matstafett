"""
Project: Food Relay
Author: Patrik Viljehav
Version: 1.0
URL: https://github.com/pvilje/matstafett
Description:    Select an Excel (.xlsx) or notepad (.txt) document.
                The script will calculate a setup where all participants
                will:
                    * Get a certain part of the dinner to prepare.
                    * Meet new people at every part of the dinner
                    * Save the result in the same file type as was submitted to the script (this can be changed)

Non standard Dependencies:
                * openpyxl verified to work on versions: 2.4.0 - 2.4.6

"""

# Import needed modules
import openpyxl
import os
import tkinter
import random
import csv
from tkinter import filedialog
from tkinter import messagebox
from openpyxl.styles import *

# Constants
# =========
OPENPYXL_VERSION = "2.4"


class Hmi:
    def __init__(self, parent, language="eng", csv_file="lang.csv"):
        """
        Initialize the class
        :param parent: The master tk window
        :param language: the GUI language, defaults to english.
        :param csv_file: The csv file containing the language strings
        """
        # Set up variables
        # ================
        self.gui_language = ""
        self.csv_file = ""
        self.file_type = ""
        self.file_name = None
        self.file_path = None
        self.list_supported_file_types = [("Excel", "*.xlsx"), ("Text", "*.txt")]
        self.list_participants = []
        self.list_sorted_participants = []
        self.groups_starter = []
        self.groups_main = []
        self.groups_desert = []
        self.list_rand_index = []
        self.num_groups = None
        self.lang = {}
        self.host_s = []
        self.host_m = []
        self.host_d = []
        self.guest_s_1 = []
        self.guest_s_2 = []
        self.guest_m_1 = []
        self.guest_m_2 = []
        self.guest_d_1 = []
        self.guest_d_2 = []
        self.prev_starter_hosts = []
        self.prev_main_hosts = []
        self.prev_desert_hosts = []

        # Get language pack
        # =================
        # Make sure it is a csv file. Allow different formats, but show a warning that it seems erroneous
        if not csv_file.lower().endswith(".csv"):
            if not messagebox.askyesno("Invalid language file",
                                       "Invalid File format for language file, expected '.csv', got .{}\n\n"
                                       "Do you want to try to use the supplied file anyway?\n"
                                       "({})".format(
                                           csv_file.lower().split(".")[-1], csv_file)):
                quit()
        # All is fine, get the phrases for the selected language
        self.csv_file = csv_file
        self.gui_language = language
        self.get_lang()
        parent.title(self.lang["title"])

        # ================
        # tkinter stuff...
        # ================

        # Main frames and labels
        # ======================
        self.l_title = tkinter.Label(parent, text=self.lang["title"])
        self.f_input = tkinter.Frame(parent, pady=10, padx=10)
        self.f_output = tkinter.Frame(parent, pady=10, padx=10)

        # TK Variables
        # ============
        self.sv_filename = tkinter.StringVar()
        self.iv_new_year_same_lineup = tkinter.IntVar()

        # Input widgets
        # =============
        self.e_filename = tkinter.Entry(self.f_input,
                                        textvariable=self.sv_filename,
                                        width=60)
        self.b_select_file = tkinter.Button(self.f_input,
                                            text=self.lang["file_select"],
                                            command=self.select_file)
        self.b_run = tkinter.Button(self.f_input,
                                    text=self.lang["button_run"],
                                    command=self.generate_result,
                                    state=tkinter.DISABLED,
                                    height=3, width=10)
        self.cb_new_year_same_lineup = tkinter.Checkbutton(self.f_input,
                                                           text=self.lang["button_same_lineup"],
                                                           variable=self.iv_new_year_same_lineup)

        # Output widgets
        # ==============
        self.t_output = tkinter.Text(self.f_output,
                                     height=10, width=70,
                                     state=tkinter.DISABLED)

        # Add colors and scroll-bars to the output widget
        # ===============================================
        log_colors = ["black", "green", "blue", "red"]
        for color in log_colors:
            self.t_output.tag_config(color, foreground=color)
        self.scroll_x_output = tkinter.Scrollbar(self.f_output, command=self.t_output.xview, orient=tkinter.HORIZONTAL)
        self.scroll_y_output = tkinter.Scrollbar(self.f_output, command=self.t_output.yview)
        self.t_output.configure(yscrollcommand=self.scroll_y_output.set, xscrollcommand=self.scroll_x_output.set)

        # Check openpyxl version
        # ======================
        # TODO, translate this.
        if not openpyxl.__version__.startswith(OPENPYXL_VERSION):
            tkinter.messagebox.showwarning("Unexpected version difference",
                                           "This program uses openpyxl version: {} \nopenpyxl version installed: {}"
                                           .format(OPENPYXL_VERSION, openpyxl.__version__))

    def draw_main(self):
        """
        Draw the main window
        """
        # Main title
        self.l_title.grid(row=0, column=0)

        # input frame
        self.f_input.grid(row=1, column=0)
        self.e_filename.grid(row=0, column=0, columnspan=2)
        self.b_select_file.grid(row=0, column=2, sticky=tkinter.W, padx=15)
        self.cb_new_year_same_lineup.grid(row=1, column=0, sticky=tkinter.W)
        self.b_run.grid(row=1, column=3)

        # output frame
        self.f_output.grid(row=2, column=0)
        self.t_output.grid(row=0, column=0, columnspan=2)
        self.scroll_x_output.grid(row=1, column=0, columnspan=2, sticky=tkinter.E + tkinter.W)
        self.scroll_y_output.grid(row=0, column=2, sticky=tkinter.N + tkinter.S)

    def select_file(self):
        """
        Method to open a file dialog and validate the file type.
        """
        # Select file dialog.
        file = filedialog.askopenfilename(title=self.lang["dialog_select_file"],
                                          initialdir=os.curdir,
                                          filetypes=self.list_supported_file_types)
        # populate String-var and entry.
        self.sv_filename.set(file)
        self.e_filename = self.sv_filename.get()

        # Validate file
        # =============
        file_type_ok = False
        if len(file) > 0:
            if file.endswith(".txt"):
                self.file_type = ".txt"
                file_type_ok = True
            elif file.endswith(".xlsx"):
                self.file_type = ".xlsx"
                file_type_ok = True
            else:
                self.log_output(self.lang["error_file_types"], "red")
        else:
            self.log_output(self.lang["error_no_file_selected"], "red")

        # Activate / deactivate run-button depending on file validation result
        # ====================================================================
        if file_type_ok:
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
        # prepare to get all participants
        self.list_participants = []
        file = os.path.join(self.file_path, self.file_name)
        self.log_output(self.lang["file_reading_{}".format(self.file_type)])

        # If it is a txt file, just loop through all the lines, assume every line is a participant
        # ========================================================================================
        if self.file_type == ".txt":
            with open(file, "r") as f:
                for line in f:
                    # Ignore blank lines
                    if line.strip():
                        self.list_participants.append(line)

        # If it is a xlsx file check column A for participants.
        # =====================================================
        elif self.file_type == ".xlsx":
            wb = openpyxl.load_workbook(file)
            ws = wb.get_sheet_by_name(wb.sheetnames[0])
            max_rows = ws.max_row
            for row in ws["A1:A{}".format(max_rows)]:
                for cell in row:
                    if cell.value is not None:
                        self.list_participants.append(cell.value)
        else:
            # Should not be possible to end up here, but just in case...
            self.log_output(self.lang["error_file_types"], "red")

    def create_lineup(self):
        """
        Create a lineup for the food rally.
        """
        # Reset variables
        # ===============
        base_index = 0
        offset_1 = 1
        offset_2 = 2
        self.host_s = []
        self.host_m = []
        self.host_d = []
        self.guest_s_1 = []
        self.guest_s_2 = []
        self.guest_m_1 = []
        self.guest_m_2 = []
        self.guest_d_1 = []
        self.guest_d_2 = []

        # loop through all groups, using three indexes for the group lists
        # =================================================================
        while base_index < self.num_groups:
            if offset_1 >= self.num_groups:
                offset_1 = 0
            if offset_2 >= self.num_groups:
                offset_2 = 0
            # starters
            self.host_s.append(self.groups_starter[base_index])
            self.guest_s_1.append(self.groups_main[base_index])
            self.guest_s_2.append(self.groups_desert[base_index])

            # main course
            self.host_m.append(self.groups_main[offset_1])
            self.guest_m_1.append(self.groups_starter[base_index])
            self.guest_m_2.append(self.groups_desert[offset_2])

            # desert
            self.host_d.append(self.groups_desert[offset_1])
            self.guest_d_1.append(self.groups_starter[base_index])
            self.guest_d_2.append(self.groups_main[offset_2])

            base_index += 1
            offset_1 += 1
            offset_2 += 1

    def save_to_file(self):
        """
        Save the generated list to a file, Grouped and neatly ordered
        """
        self.log_output(self.lang["progress_gen_route"])

        # Save a .txt file
        # ================
        if self.file_type == ".txt":

            # Create a new file since we don't want to mess with the source.
            filename = "new_" + self.file_name
            # result = starter + main + desert
            # Open / create a new file and save the results.
            with open(os.path.join(self.file_path, filename), "w", encoding="utf8") as f:
                f.write("{}\n".format(self.lang["starter"]))
                for index, host in enumerate(self.host_s):
                    f.write("{}: {}".format(self.lang["host"], host))
                    f.write("{}: {}".format(self.lang["guest"], self.guest_s_1[index]))
                    f.write("{}: {}\n".format(self.lang["guest"], self.guest_s_2[index]))

                f.write("{}\n".format(self.lang["main_course"]))
                for index, host in enumerate(self.host_m):
                    f.write("{}: {}".format(self.lang["host"], host))
                    f.write("{}: {}".format(self.lang["guest"], self.guest_m_1[index]))
                    f.write("{}: {}\n".format(self.lang["guest"], self.guest_m_2[index]))

                f.write("{}\n".format(self.lang["desert"]))
                for index, host in enumerate(self.host_d):
                    f.write("{}: {}".format(self.lang["host"], host))
                    f.write("{}: {}".format(self.lang["guest"], self.guest_d_1[index]))
                    f.write("{}: {}\n".format(self.lang["guest"], self.guest_d_2[index]))

            self.log_output("{} \n{}".format(self.lang["progress_saved_to"], filename))

        # Save a .xlsx file
        # =================
        elif self.file_type == ".xlsx":
            # Save to a new file, don't mess with the source.
            # check if this is the first generated result
            if os.path.isfile(os.path.join(self.file_path, "{}_{}".format(self.lang["result"], self.file_name))):
                # File already existed, generate a new filename (add a number until a not used name is found)
                file_no = 2
                while os.path.isfile(
                        os.path.join(
                            self.file_path, "{}_{}_{}".format(self.lang["result"], str(file_no), self.file_name))):
                    file_no += 1
                file = os.path.join(
                    self.file_path, "{}_{}_{}".format(self.lang["result"], str(file_no), self.file_name))
            else:
                # This is the first generated result.
                file = os.path.join(self.file_path, "{}_{}".format(self.lang["result"], self.file_name))

            # Open Excel workbook
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = self.lang["title"]

            # Setup styles
            # ============
            ws.column_dimensions["A"].width = 34
            ws.column_dimensions["C"].width = 34
            ws.column_dimensions["D"].width = 34
            ws.column_dimensions["E"].width = 34
            h1 = NamedStyle(name="h1", font=Font(size=15, bold=True, color="1f497d"),
                            border=Border(bottom=Side(style="thick", color="4f81bd")))
            h1_center = NamedStyle(name="h1_center", font=Font(size=15, bold=True, color="1f497d"),
                                   border=Border(bottom=Side(style="thick", color="4f81bd")),
                                   alignment=Alignment(horizontal="center"))
            h2 = NamedStyle(name="h2", font=Font(size=13, bold=True, color="1f497d"),
                            border=Border(bottom=Side(style="thick", color="a7bfde")))
            h2_center = NamedStyle(name="h2_center", font=Font(size=13, bold=True, color="1f497d"),
                                   border=Border(bottom=Side(style="thick", color="a7bfde")),
                                   alignment=Alignment(horizontal="center"))
            # Column A, summary of participants
            # =================================

            # Summary header
            ws["A1"] = "{}".format(self.lang["summary"])
            ws["A1"].style = h1

            # Starter hosts
            ws["A2"] = "{} {}:".format(self.lang["host"], self.lang["starter"])
            ws["A2"].style = h2
            row = 3
            for name in self.host_s:
                ws["A{}".format(row)] = name
                row += 1
            row += 2

            # Main course hosts
            ws["A{}".format(row)] = "{} {}:".format(self.lang["host"], self.lang["main_course"])
            ws["A{}".format(row)].style = h2
            row += 1
            for name in self.host_m:
                ws["A{}".format(row)] = name
                row += 1
            row += 2

            # Desert hosts
            ws["A{}".format(row)] = "{} {}:".format(self.lang["host"], self.lang["desert"])
            ws["A{}".format(row)].style = h2
            row += 1
            for name in self.host_d:
                ws["A{}".format(row)] = name
                row += 1

            # Column C, D, E: results
            # =================================
            # Starters
            ws.merge_cells("C1:E1")
            ws["C1"] = self.lang["starter"]
            ws["C1"].style = h1_center
            ws["D1"].style = h1_center  # needed for the border, even though the cells are merged
            ws["E1"].style = h1_center  # needed for the border, even though the cells are merged
            ws["C2"] = self.lang["host"]
            ws["D2"] = self.lang["guest"]
            ws["E2"] = self.lang["guest"]
            ws["C2"].style = h2_center
            ws["D2"].style = h2_center
            ws["E2"].style = h2_center
            row = 3
            for index, host in enumerate(self.host_s):
                ws["C{}".format(row)] = host
                ws["D{}".format(row)] = self.guest_s_1[index]
                ws["E{}".format(row)] = self.guest_s_2[index]
                row += 1

            # Main Course
            row += 1
            ws.merge_cells("C{}:E{}".format(row, row))
            ws["C{}".format(row)] = self.lang["main_course"]
            ws["C{}".format(row)].style = h1_center
            ws["D{}".format(row)].style = h1_center  # needed for the border, even though the cells are merged
            ws["E{}".format(row)].style = h1_center  # needed for the border, even though the cells are merged
            row += 1
            ws["C{}".format(row)] = self.lang["host"]
            ws["D{}".format(row)] = self.lang["guest"]
            ws["E{}".format(row)] = self.lang["guest"]
            ws["C{}".format(row)].style = h2_center
            ws["D{}".format(row)].style = h2_center
            ws["E{}".format(row)].style = h2_center
            row += 1
            for index, host in enumerate(self.host_m):
                ws["C{}".format(row)] = host
                ws["D{}".format(row)] = self.guest_m_1[index]
                ws["E{}".format(row)] = self.guest_m_2[index]
                row += 1

            # Desert
            row += 1
            ws.merge_cells("C{}:E{}".format(row, row))
            ws["C{}".format(row)] = self.lang["desert"]
            ws["C{}".format(row)].style = h1_center
            ws["D{}".format(row)].style = h1_center  # needed for the border, even though the cells are merged
            ws["E{}".format(row)].style = h1_center  # needed for the border, even though the cells are merged
            row += 1
            ws["C{}".format(row)] = self.lang["host"]
            ws["D{}".format(row)] = self.lang["guest"]
            ws["E{}".format(row)] = self.lang["guest"]
            ws["C{}".format(row)].style = h2_center
            ws["D{}".format(row)].style = h2_center
            ws["E{}".format(row)].style = h2_center
            row += 1
            for index, host in enumerate(self.host_d):
                ws["C{}".format(row)] = host
                ws["D{}".format(row)] = self.guest_d_1[index]
                ws["E{}".format(row)] = self.guest_d_2[index]
                row += 1

            # Save! (also closes the file)
            try:
                wb.save(filename=file)
            except PermissionError:
                self.log_output(self.lang["error_save"], "red")
            except FileNotFoundError:
                self.log_output(self.lang["error_save"], "red")
            else:
                self.log_output(self.lang["progress_done"])
                self.log_output("{} \n{}".format(self.lang["progress_saved_to"], file))
                messagebox.showinfo(self.lang["progress_done"], "{}: {}".format(self.lang["dialog_done_msg"], file))

        else:
            # Should not be possible to end up here, but just in case...
            self.log_output(self.lang["error_file_types"], "red")

    def validate_number_of_participants(self):
        """
        Makes sure the number of participants is:
            * a factor of 3.
            * at least 9
        Raises ValueError if not.
        """
        # Get all the participants:
        # =========================
        # Is this based on a previous result or not?
        if self.iv_new_year_same_lineup.get():
            list_participants = self.prev_starter_hosts + self.prev_main_hosts + self.prev_desert_hosts
            # Make sure the three groups are equal in size
            if not (len(self.prev_starter_hosts) == len(self.prev_main_hosts) and
                    len(self.prev_main_hosts) == len(self.prev_desert_hosts)):
                # Todo: Translate this.
                raise ValueError("The previous starter,main and desert hosts are not distributed equally")
        else:
            list_participants = self.list_participants

        # Verify that it is an ok number of participants
        # ==============================================
        if len(list_participants) < 9:
            # Must be 9 or more
            raise ValueError(self.lang["error_less_than_nine"])
        elif len(list_participants) % 3 != 0:
            # Must be a factor of 3
            raise ValueError(self.lang["error_number_participants"])
        else:
            # All is ok, go on.
            self.num_groups = int(len(list_participants) / 3)
            self.log_output("{}: {}".format(self.lang["progress_found_participants"], len(list_participants)))

    def generate_random_index(self):
        """
        Generate a list containing integers from 1 - amount of participants.
        Then shuffle it around so they don't all come in order...
        """
        self.log_output(self.lang["progress_gen_rand_list"])
        i = 0
        self.list_rand_index = []
        while i < len(self.list_participants):
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
            self.list_sorted_participants.append(self.list_participants[index])

        # create three equal sized lists containing all participants
        # ==========================================================
        self.groups_starter = self.list_sorted_participants[0:self.num_groups]
        self.groups_main = self.list_sorted_participants[self.num_groups:self.num_groups * 2]
        self.groups_desert = self.list_sorted_participants[self.num_groups * 2:self.num_groups * 3]

    def get_previous_lineup(self):
        """
        Get the list of previous participants.
        :return result: 0 - OK, 1 - Unsupported File, 2 - Faulty DataBool
        """
        # This will only work for excel files.
        if self.file_type != ".xlsx":
            return 1
        else:
            # Reset needed variables
            # ======================
            languages = []
            host = []
            self.prev_starter_hosts = []
            self.prev_main_hosts = []
            self.prev_desert_hosts = []
            get_participants = [False, "starter", "main", "desert"]  # which type of participants to get.
            p_get_participants = 0  # pointer to get_participants
            excel_file = os.path.join(self.file_path, self.file_name)  # save the excel file path

            # get some phrases we need in all possible languages.
            # ===================================================
            with open(self.csv_file, "r", encoding="utf8") as csv_file:
                reader = csv.DictReader(csv_file, delimiter=",")

                # get all available languages.
                # ===========================
                # first get all headers.
                for header in reader.fieldnames:
                    languages.append(header)
                # then remove the first header, it is not a language
                languages.pop(0)
                # Now get the "host" phrase in all known/supported languages.
                for row in reader:
                    if row["phrase"] == "host":
                        for language in languages:
                            # add a space after the text to reduce the risk of the string being a part of a name
                            host.append("{} ".format(row[language]))

            # Open the excel file and search for participants
            # ===============================================
            wb = openpyxl.load_workbook(excel_file)
            ws = wb.get_sheet_by_name(wb.sheetnames[0])
            max_rows = ws.max_row
            for row in ws["A1:A{}".format(max_rows)]:
                for cell in row:
                    if cell.value is not None:
                        # skip blank lines.
                        if cell.value == "":
                            pass
                        # Check for headlines
                        elif any(word in cell.value for word in host):
                            p_get_participants += 1  # increase the pointer
                        # Add the previous starters hosts
                        elif get_participants[p_get_participants] == "starter":
                            self.prev_starter_hosts.append(cell.value)
                        # Add the previous main course hosts
                        elif get_participants[p_get_participants] == "main":
                            self.prev_main_hosts.append(cell.value)
                        # Add the previous desert hosts
                        elif get_participants[p_get_participants] == "desert":
                            self.prev_desert_hosts.append(cell.value)

            # Quick check, if this has worked p_get_participants show now have reached max.
            if not p_get_participants == len(get_participants) - 1:
                return 2
            # Success!
            return 0

    def generate_result(self):
        """
        Start the process of generating the results.
        """
        # See if previous line up should be taken into account.
        if self.iv_new_year_same_lineup.get():
            # Try to get the previous participants.
            previous_line_up_read = self.get_previous_lineup()
            if previous_line_up_read == 0:
                # Todo, translate this
                self.log_output("Previous lineup read.")
            elif previous_line_up_read == 1:
                # Todo, translate this
                self.log_output("Unsupported file (needs to be .xlsx).", "red")
                return
            elif previous_line_up_read == 2:
                # Todo, translate this
                self.log_output("Invalid data in excel file. Cannot complete a new lineup.", "red")
                return
            else:
                # Todo, translate this
                self.log_output("Unexpected error when reading the file, no idea what went wrong", "red")
                return

        else:
            # Just read the new file contents.
            self.read_file_contents()
        try:
            self.validate_number_of_participants()
        except ValueError as e:
            self.log_output("{}: {}".format(self.lang["error"], e), "red")
            return

        # Only create new groups of participants if previous line up not should be taken into account.
        if not self.iv_new_year_same_lineup.get():
            self.generate_random_index()
            self.sort_participants()

        # Generate a new lineup.
        self.create_lineup()  # Todo this needs to be fixed for new_year_same_lineup

        self.log_output(self.lang["progress_done_saving"])

        # Save the result
        self.save_to_file()

    def log_output(self, text, color="black"):
        """
        Method to print text to the output frame.
        :param text: The text to print
        :param color: Text color
        """
        self.t_output.configure(state=tkinter.NORMAL)
        if not text.endswith("\n"):
            text += "\n"
        self.t_output.insert(tkinter.END, text, color)
        self.t_output.yview_moveto(1)
        self.t_output.configure(state=tkinter.DISABLED)

    def get_lang(self):
        """
        Read a csv file and return the phrases matching the selected language
        self.lang is set to a dict with phrases in selected language
        """
        cur_lang = {}
        try:
            with open(self.csv_file, "r", encoding="utf8") as csv_file:
                reader = csv.DictReader(csv_file, delimiter=",")
                for row in reader:
                    cur_lang[row["phrase"]] = row[self.gui_language]
        except KeyError:
            pass
        except FileNotFoundError:
            messagebox.showerror("File not found", "The language file: '{}' can not be found.\n"
                                 "Make sure it is available and try again.".format(self.csv_file))
            quit()
        self.lang = cur_lang

        # ============
        # END OF CLASS
        # ============

if __name__ == "__main__":
    root = tkinter.Tk()
    hmi = Hmi(root, language="swe")
    hmi.draw_main()
    root.mainloop()
