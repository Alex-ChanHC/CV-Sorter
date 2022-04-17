if __name__ != '__main__':
    import tkinter as tk
    import tkinter.filedialog as tkFiledialog
    import tkinter.messagebox as tkMessageBox
    import functions

    root = tk.Tk()

    class UI():

        Input_Keys = (
            "ent_dir_xls",
            "ent_filename",
            "om_filetype",
            "ent_row_no",
            "ent_col_UID",
            "ent_col_nom",
            "cbtn_condi_mssheet",
            "ent_condi_mssheet",
            "ent_condi_col_post",
            "ent_dir_cv",
            "cbtn_multi_cv_fld",
            "ent_dir_des",
            "cbtn_seg_cv",
            "cbtn_prefix"
        )

        Labels = (
            "Direcotry of Excel:",
            "Excel File Name:",
            "Excel File Type:",
            "Row No. of Header Row:",
            "Column Heading for UID:",
            "Column Heading for Nomination Status:",
            "All sheets combined to 1 Mastersheet?",
            "Name of Mastersheet:",
            "Column Heading for Job Post ID:",
            "Directory of CV:",
            "Multiple Folders of CVs?",
            "Place Nominated CVs in:",
            "Segregate CVs into Multiple Folders?",
            "Add Prefix (Job Code) for CVs?"
            )

        UserInput = { key:
                {
                    "lbl":  None,
                    "obj":  None,
                    "var":  None,
                    "val":  None
                }
                for key in Input_Keys
            }

        for i, key in enumerate(Input_Keys):
            UserInput[key]["lbl"] = Labels[i]
            if "cbtn" in key:
                UserInput[key]["var"] = tk.IntVar()
            if "om" in key:
                UserInput[key]["var"] = tk.StringVar()

        Filetype = [
            ".xlsx",
            ".xls"
        ]

        UserInput["om_filetype"]["var"].set(Filetype[0])

        def __init__(self, master):
            self.master = master
            master.title("CV Sorter 2.0")

            upperFrm = tk.Frame(
                master,
                relief = tk.SUNKEN,
                borderwidth = 5
                )
            upperFrm.grid(row = 0, column = 0, sticky = "nesw", padx = 5, pady = (10,5))

            i = 0
            while i < len(self.Input_Keys):
                current_key = self.Input_Keys[i]

                #Create descriptive texts for fields
                self.lbl = tk.Label(
                    upperFrm,
                    text = self.Labels[i],
                    font = ("Calibri", 12)
                    ).grid(
                        row = i,
                        column = 0,
                        sticky = "w"
                        )

                #Create input fields
                #Can upgrade to function with guard clause
                if "ent" in current_key:
                    self.UserInput[current_key]["obj"] = tk.Entry(
                        upperFrm,
                        width = 70
                        )
                    if "condi" in current_key:
                        self.UserInput[current_key]["obj"].configure(
                            state = "disabled"
                        )

                if "cbtn" in current_key:
                    self.UserInput[current_key]["obj"] = tk.Checkbutton(
                        upperFrm,
                        text = "Yes",
                        variable = self.UserInput[current_key]["var"]
                        )
                    if "condi" in current_key:
                        self.UserInput[current_key]["obj"].configure(
                                command = lambda i = current_key:
                                    self.activate(i)
                            )

                if "om_" in current_key:
                    self.UserInput[current_key]["obj"] = tk.OptionMenu(
                        upperFrm,
                        self.UserInput[current_key]["var"],
                        *self.Filetype
                    )

                self.UserInput[current_key]["obj"].grid(row = i, column = 1)

                if "dir" in current_key:
                    locals()[f"self.btn_browse_{current_key}"] = tk.Button(
                        upperFrm,
                        text = "Browse",
                        command = lambda current_key = current_key :
                            self.get_dir(current_key)
                    ).grid(
                        row = i,
                        column = 2,
                    )

                i += 1

            lowerFrm = tk.Frame(
                master,
                )
            lowerFrm.grid(row = 1, column = 0, sticky = "nesw", padx = 5, pady = 5)

            self.btn_table = tk.Button(
                lowerFrm,
                text = "Preview Table",
                command = lambda: self.preview_table()
            ).grid(
                row = 1,
                column = 0
            )

            self.btn_sort = tk.Button(
                lowerFrm,
                text = "Begin Sort",
                command = lambda: self.sort()
            ).grid(
                row = 1,
                column = 1
            )


        def get_dir(self, key):
            if "xls" in key:
                filepath = tkFiledialog.askopenfilename(
                    filetype = [("Excel Files", "*.xlsx *.xls")]
                )
            else:
                filepath = tkFiledialog.askdirectory()

            if not filepath:
                return False

            if "xls" in key:
                self.set_text(
                    self.UserInput[key]["obj"],
                    filepath[:filepath.rfind("/")]
                    )
                self.set_text(
                    self.UserInput["ent_filename"]["obj"],
                    filepath[filepath.rfind("/") + 1:filepath.rfind(".")]
                    )
                self.UserInput["om_filetype"]["var"].set(
                    filepath[filepath.rfind("."):]
                    )
                return 1

            self.set_text(
                self.UserInput[key]["obj"],
                filepath
            )
            return 2

        def set_text(self, obj, text):
            obj.delete(0, tk.END)
            obj.insert(0, text)
            return True

        def activate(self, key):
            elements = ["ent_condi_mssheet", "ent_condi_col_post"]
            if self.UserInput[key]["var"].get():
                for i in elements:
                    self.UserInput[i]["obj"].configure(
                        state = "normal"
                    )
                return True
            for i in elements:
                self.UserInput[i]["obj"].configure(
                    state = "disabled"
                )
            return False

        def sync_input(self):
            for key in self.Input_Keys:
                if "ent" in key:
                    self.UserInput[key]["val"] = self.UserInput[key]["obj"].get()
                else:
                    self.UserInput[key]["val"] = self.UserInput[key]["var"].get()
            
            """Insert try except validation block"""
            #return False
            return True

        def preview_table(self):
            if self.sync_input():
                dummyList = [innerdict["val"] for _, innerdict in self.UserInput.items()]
                # dummyList = ['C:/Users/alexc/Downloads', 'Applicant list_CEDARS_Shirley', '.xlsx', '1', 'U №', 'Nomination', 1, 'Final', 'Job Ref No.', '', 0, '', 0, 0]
                fileObj = functions.NominationObj(*dummyList)
                fileObj.read_excel()
                popUp = tk.Toplevel(root)
                popUp.title("Preview Table")
                tk.Label(
                    popUp,
                    text = fileObj.preview_dataframe(),
                    font = ("Calibri", 12)
                ).pack(padx = 10, pady = 10)
                return True
            return False
        
        def sort(self):
            if self.sync_input():
                dummyList = [innerdict["val"] for _, innerdict in self.UserInput.items()]
                fileObj = functions.NominationObj(*dummyList)
                fileObj.read_excel()
                fileObj.nominate_func()
                return True
            return False

        @staticmethod
        def _show_error(msg):
            tkMessageBox.showerror("Error", msg + "\n" + "Please seek help from Alex")
            return



    root.minsize(755, 420)
    myUI = UI(root)
    #screen_width, screen_height  = root.winfo_width(), root.winfo_height()
    #Print the screen size
    #print(f"Screen width: {screen_width}", f"Screen height: {screen_height}", "\n")
