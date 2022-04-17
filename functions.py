if __name__ != '__main__':
    import os, os.path, shutil
    import pandas as pd


    def decorative_func(func):
        def wrap():
            print('\n==========')
            func()
            print('==========\n')
        return wrap()

    class NominationObj():
        keys = (
            "dir_xls",
            "filename",
            "filetype",
            "row",
            "uid",
            "nom",
            "consolidated",
            "sheetname",
            "post_id",
            "dir_cv",
            "multi_cv_fld",
            "dir_des",
            "seg_cv",
            "prefix"
        )

        __match_choices = ["gp == fld", "gp in fld"]
        __match_condition = __match_choices[0]

        def __init__(self, *args):
            attr = { self.keys[i] : args[i] for i in range(len(args)) }
            self.attr = attr
            self.df = None
            self.df_gp = None

        def read_excel(self):
            dir_xls = self.attr["dir_xls"]
            filename = self.attr["filename"]
            filetype = self.attr["filetype"]
            row = self.attr["row"]

            if not self.attr["consolidated"]:
                self.attr["sheetname"] = None
                self.attr["post_id"] = None

                raw_data = pd.read_excel(
                  f"{dir_xls}/{filename}{filetype}",
                  skiprows = int(row)-1,
                  sheet_name = None
                  )

                data = pd.DataFrame()
                for name, sheet in raw_data.items():
                    sheet['Worksheet'] = name
                    data = data.append(sheet)
                self.attr["post_id"] = 'Worksheet'
                '''Assuming Worksheet name is Post ID
                Forgot what that meant'''

            else:
                data = pd.read_excel(
                  f"{dir_xls}/{filename}{filetype}",
                  converters = { self.attr["post_id"] : str },
                  skiprows = int(row)-1,
                  sheet_name = self.attr["sheetname"]
                  )

            relevant_data = pd.DataFrame(
                data,
                columns = [
                    self.attr["uid"],
                    self.attr["post_id"],
                    self.attr["nom"]
                    ]
                )

            filtered_table = relevant_data[
                relevant_data[self.attr["nom"]].fillna('') != ''
                ]

            self.df = filtered_table.set_index(self.attr["uid"])
            self.df_gp = self.df.groupby(self.attr["post_id"])
            return True

        def preview_dataframe(self):
            Text = [ 
                'Preview of the first 10 rows: ',
                self.df.head(10).to_string(col_space = 20),
                "Size of filtered Table: [{m} rows x {n} columns]".format(
                    m = self.df.shape[0] ,
                    n = self.df.shape[1]
                    )
            ]
            return "\n\n".join(Text)

        @classmethod
        def _set_match_condition(cls, value):
            cls.__match_condition = cls.__match_choices[value]
            return True

        def add_prefix_func(self, filename, prefix, indicator):
            if indicator:
                return prefix + filename
            return filename

        def separate_folders(self, filename, path_Y, path_N, indicator):
            if indicator:
                return path_Y + '/' + filename
            return path_N + '/' + filename

        def nominate_func(self):
            if self.attr["multi_cv_fld"]:
                self.nominate_multi()
                return 1
            self.nominate_single()
            return 0

        def _write_error_txt(self, error_count, error_dict, directory):
            if error_count:
                with open(directory+ "Error.txt", "w") as err:
                    for k, v in error_dict.items():
                        if v != []:
                            err.write('{0} : {1}\n'.format(k, v))
                    err.close()
                def inter():
                    print(
                        f"No. of files failed to be read (Unexpected name format): {error_count}",
                        f"Corresponding errors can be found in:\n {directory}\n",
                    sep = "\n"
                    )
                    for k, v in error_dict.items():
                        if v != []:
                            print(k ," : ", v)
                decorative_func(inter)        
            return True

        def nominate_multi(self):
            dir_fld = self.attr["dir_cv"] + "/"
            dir_des = self.attr['dir_des']
            segregate = self.attr['seg_cv']
            add_prefix = self.attr['prefix']
            
            flds = [fld for fld in os.listdir(dir_fld) if os.path.isdir(dir_fld + fld)]
            errFiles, errCount = {}, 0
            groups = [gp for gp,_ in self.df_gp]
            for gp in groups:
                df_prj = self.df_gp.get_group(gp)
                Nom_UID_list = list(df_prj.index.values)
                for fld in flds:
                    if eval(self.__match_condition):
                        errFiles[fld] = []
                        print("Currently sorting CV folder: %s" % fld)
                        path = os.path.join(dir_des, gp)
                        if segregate and (not os.path.exists(path)):
                            os.mkdir(path)
                        for file in os.listdir(dir_fld + fld):
                            if file.endswith((".pdf" , ".doc" , ".docx")):
                                try:
                                    """Can Regex the uid"""
                                    uid = int(file[file.index("_") + 1 :file.find(".")])
                                    if uid in Nom_UID_list:
                                        """Refactor this part"""
                                        Neuer_fName = file[:file.index("_")] + file[file.find("."):]
                                        prefix = df_prj.loc[int(uid)][self.attr["post_id"]] + '_'
                                        res = self.separate_folders(
                                        self.add_prefix_func(
                                            Neuer_fName, prefix, add_prefix
                                            ),
                                        path, dir_des, segregate
                                        )
                                        # shutil.copyfile(f"\\\\?\\{dir_fld}\\{file}" , res)
                                        normpath = os.path.normpath(f"{dir_fld + fld}\\{file}")
                                        shutil.copyfile(u'\\\\?\\{}'.format(normpath) , res)                                       
                                except ValueError:
                                    errFiles[fld].append(file)
                                    errCount += 1
                                    pass

            self._write_error_txt(errCount, errFiles, dir_des)
            return True
        
        def nominate_single(self):
            dir_fld = self.attr["dir_cv"] + "/"
            dir_des = self.attr['dir_des']
            segregate = self.attr['seg_cv']
            add_prefix = self.attr['prefix']

            errFiles, errCount = {}, 0
            groups = [gp for gp,_ in self.df_gp]
            for gp in groups:
                df_prj = self.df_gp.get_group(gp)
                Nom_UID_list = list(df_prj.index.values)
                errFiles[dir_fld] = []
                print("Currently sorting CV folder: %s" % dir_fld)
                path = os.path.join(dir_des, gp)
                if segregate and (not os.path.exists(path)):
                    os.mkdir(path)
                for file in os.listdir(dir_fld):
                    if file.endswith((".pdf" , ".doc" , ".docx")):
                        try:
                            uid = int(file[file.index("_") + 1 :file.find(".")])
                            if uid in Nom_UID_list:
                                Neuer_fName = file[:file.index("_")] + file[file.find("."):]
                                prefix = df_prj.loc[int(uid)][self.attr["post_id"]] + '_'
                                res = self.separate_folders(
                                self.add_prefix_func(
                                    Neuer_fName, prefix, add_prefix
                                    ),
                                path, dir_des, segregate
                                )
                                normpath = os.path.normpath(f"{dir_fld}\\{file}")
                                shutil.copyfile(u'\\\\?\\{}'.format(normpath) , res)  
                        except ValueError:
                            errFiles[dir_fld].append(file)
                            errCount += 1
                            pass

            self._write_error_txt(errCount, errFiles, dir_des)
            return True