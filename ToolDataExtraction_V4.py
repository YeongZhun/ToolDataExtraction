# -*- coding: utf-8 -*-
"""
Created on Sun Nov 20 21:17:50 2022

@author: TMR
"""

import numpy as np
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell
from xlsxwriter.utility import xl_cell_to_rowcol
import glob
import os
import pandas as pd
from pandas import ExcelWriter
import tkinter as tk
from tkinter import *
from tkinter import ttk, filedialog, scrolledtext, StringVar, Label, Button, Entry
from tkinter.scrolledtext import ScrolledText
from pathlib import Path
import xlwings as xw
import threading
from threading import Thread
from statistics import mean
import copy

global countdies
countdies = []

class OT_Frame(ttk.Frame):
    def __init__(self, container):
        super().__init__(container)

        self.columnconfigure(0, weight=1)
        self.columnconfigure(1, weight=3)
        
        self.create_OT_widgets()
        
    def create_OT_widgets(self):
        
        self.label = Label(self, text="This is for Optical Test, please select the FOLDER.", background='#FFCCE5', width=80)
        self.label.grid(row=0, column=0, columnspan=3, sticky=tk.EW, padx=5, pady=5)
        
        OT_Entry_Rawfileloc = StringVar(self, value='Raw Data Location')
        self.OT_Entry_Raw = Entry(self, width=50, textvariable=OT_Entry_Rawfileloc, foreground='gray')
        self.OT_Entry_Raw.grid(row=1, column=1,columnspan=1, sticky=tk.W, padx=1, pady=1)
        
        self.button1 = Button(self,
                        text="Select OT File Location:",
                        command=self.thread_OTselectdir,
                        font=("Comic Sans",11),
                        relief=GROOVE,
                        state=ACTIVE)
        self.button1.grid(row=1, column=0, sticky=W, padx=1, pady=1)
        
        self.button2 = Button(self,
                          text="Save Extracted Data to: ",
                          command=self.thread_OTsavefiles,
                          font=("Comic Sans",11),
                          relief=GROOVE,
                          state=DISABLED)
        self.button2.grid(row=2, column=0, sticky=W, padx=1, pady=1)
        
        OT_Entry_Savefileloc = StringVar(self, value='Saved Data Location')
        self.OT_Entry_Saved = Entry(self, width=50, textvariable=OT_Entry_Savefileloc, foreground='gray')
        self.OT_Entry_Saved.grid(row=2, column=1,columnspan=1, sticky=tk.W, padx=1, pady=1)
        
        self.button3 = Button(self, 
                          text="Extract Files",
                          command=self.thread_OTextractfiles,
                          font=("Comic Sans",11),
                          relief=GROOVE,
                          state=DISABLED)
        self.button3.grid(row=1, column=2, sticky=NSEW, rowspan=2, padx=1, pady=1)
        
        self.OT_sep = ttk.Separator(self,orient='horizontal')
        self.OT_sep.grid(row=4, columnspan=3, sticky='EW', pady=1)

    def thread_OTselectdir(self):
        return Thread(daemon=True, target=self.OT_SelectDir).start()

    def thread_OTsavefiles(self):
        return Thread(daemon=True, target=self.OT_Savefiles).start()

    def thread_OTextractfiles(self):
        return Thread(daemon=True, target=self.OT_ExtractFiles).start()

    def OT_Savefiles(self):
        try:
            global OT_Save_Location
            self.OT_Entry_Saved.config(state=NORMAL)
            self.OT_Entry_Saved.delete(0,'end')
            OT_Save_Location = filedialog.asksaveasfilename(filetypes = [('Excel', '*.xlsx')], initialdir=os.path.expanduser("~/Desktop"))
            self.OT_Entry_Saved.insert(tk.END, OT_Save_Location)
            self.OT_Entry_Saved.config(state=DISABLED)
            self.button3.config(state=NORMAL)
            App.thread_inserttext(app, text="Data will be saved in: "+OT_Save_Location+"\n")
            print("Data will be saved in: "+OT_Save_Location)
        except:
            App.thread_inserttext(app, text="There is an error in saving file name =(\nPlease try again.")
            pass
        
    def OT_SelectDir(self):
        try:
            App.thread_inserttext(app, text="Optical Test Selected!\n")
            App.thread_inserttext(app, text="Please start from selecting the OT Folder.\n")
            global OT_filedir
            self.OT_Entry_Raw.config(state=NORMAL)
            self.OT_Entry_Raw.delete(0,'end')
            OT_filedir = filedialog.askdirectory(initialdir="K:\\Inspection database\\3. O-Test\\O-Test Data\\2022")
            App.thread_inserttext(app, text="Selected OT Folder: "+OT_filedir+"\n")
            print("OT Dir: "+OT_filedir)
            print("\n")
            lotsdir = OT_filedir
            
            self.OT_Entry_Raw.insert(tk.END, OT_filedir)
            self.OT_Entry_Raw.config(state=DISABLED)
            
            #To get to number of slots
            for file in glob.iglob(lotsdir+"/Slot*"):
                print (file)
                count = 0
            
                #To obtain total number of dies per slot (To find out how many rows to extract in raw file later)
                for slot in glob.iglob(file+"/Die*"):
                    count+=1
                countdies.append(count)
            # print(countdies)
            self.button2.config(state=NORMAL)
        except:
            App.thread_inserttext(app, text="...\nThere is an error with selecting the OT Folder =(\nPlease check if OT Folder format has any issues.\nFolder should have SlopeMeas, and all the Slots")
            pass    
    def OT_ExtractFiles(self):
        try:
            # print(OT_filedir)
            App.thread_inserttext(app, text="OT Data is being processed...\n")
            df_per_slot = []
            combined_dataframe = pd.DataFrame()
            slopemeasure = OT_filedir+"/SlopeMeas"
            Slotlist = []
            slotnamelist = []
            slotcount = 0
            column_list = []
            
            for file in glob.iglob(slopemeasure+"/*.xlsx"):
                slotnamelist.append(file)
                # lot_name = file.split('_')[0]
                #To extract Slot ID from filename, to put them in list for naming later if necessary
                lowercase_file = file.lower()
                findslot = lowercase_file.find("slot")
                slot_id = file[findslot:findslot+6]
                if slot_id[-1].isnumeric() != True:
                    slot_id = slot_id[:-1]
                Slotlist.append(slot_id)
            
                file_pd = pd.read_excel(file, sheet_name=0, skiprows=6)
                copy_table = file_pd.iloc[0:countdies[slotcount],:].copy()
                copy_table.insert(0,"Slot ID",Slotlist[slotcount])
                copy_table.columns.values[1] = "Die No."

                # copy_table.loc[:, "Avg_Chn_Loss"] = ""
                # copy_table.loc[:, "Avg_Slab_Loss"] = ""
                # copy_table.loc[:, "Avg_SiN_Loss"] = ""
        
                # try:
                #     find_channel_loss_index = copy_table.columns.get_loc("E6(1.68)")
                #     copy_table.rename(columns={ copy_table.columns[find_channel_loss_index+1]: "Chn Loss" }, inplace = True)
                #     copy_table["Chn Loss"] = copy_table["Chn Loss"].abs()
                #     avg_chn_loss = copy_table.iloc[:,find_channel_loss_index+1].mean(skipna=True)
                #     copy_table.loc[0,"Avg_Chn_Loss"] = avg_chn_loss
                        
                #     find_slab_loss_index = copy_table.columns.get_loc("E9(1.68)")
                #     copy_table.rename(columns={ copy_table.columns[find_slab_loss_index+1]: "Slab Loss" }, inplace = True)
                #     copy_table["Slab Loss"] = copy_table["Slab Loss"].abs()
                #     avg_slab_loss = copy_table.iloc[:,find_slab_loss_index+1].mean(skipna=True)
                #     copy_table.loc[0,"Avg_Slab_Loss"] = avg_slab_loss
                    
                #     # find_SiN_loss_index = copy_table.columns.get_loc("N3(304)")
                #     # copy_table.rename(columns={ copy_table.columns[find_SiN_loss_index+1]: "SiN Loss" }, inplace = True)
                #     # copy_table["SiN Loss"] = copy_table["SiN Loss"].abs()
                #     # avg_SiN_loss = copy_table.iloc[:,find_SiN_loss_index+1].mean(skipna=True)
                #     # copy_table.loc[0,"Avg_SiN_Loss"] = avg_SiN_loss
                        
                # except:
                #     print("There is an error.")
                #     pass

                df_per_slot.append(copy_table)
                combined_dataframe = pd.concat(df_per_slot, ignore_index=True)
                slotcount+=1

            with pd.ExcelWriter(OT_Save_Location+".xlsx", engine='xlsxwriter') as writer:
                combined_dataframe.style.set_properties(**{'text-align': 'center'}).to_excel(writer, sheet_name="Optical Loss", startrow=2, startcol=2, index=False)


            # OT_excel_file = pd.ExcelWriter(OT_Save_Location+".xlsx")
            # combined_dataframe = combined_dataframe.style.set_properties(**{'text-align': 'center'})
            # combined_dataframe.to_excel(OT_excel_file, sheet_name="Optical Loss", startrow=2, startcol=2, index=False)

            # column_list = combined_dataframe.columns.values.tolist()
            # #+2 because final excel sheets starts after 2 column (and 2 row)
            # Chn_col_list = column_list.index("Avg_Chn_Loss") + 2
            # Slab_col_list = column_list.index("Avg_Slab_Loss") + 2
            # SiN_col_list = column_list.index("Avg_SiN_Loss") + 2
                (max_row, max_col) = combined_dataframe.shape
                workbook = writer.book
                worksheet = writer.sheets['Optical Loss']

                # center_align = workbook.add_format()
                # center_align.set_align('vcenter')
                bold = workbook.add_format({'bold': True})
                # worksheet.write((2, 2, max_row + 2, max_col + 1), " ", center_align)
                worksheet.write('C2', OT_filedir, bold)
                
                column_settings = [{'header': column} for column in combined_dataframe.columns]
                worksheet.add_table(2, 2, max_row + 2, max_col + 1, {'columns': column_settings, 'autofilter': False})
                worksheet.set_column(0, max_col+1, 15)
                worksheet.set_zoom(90)


    #        for sheetname, df in combined_dataframe.items():  # loop through `dict` of dataframes
    #            # combined_dataframe.to_excel(OT_excel_file, sheet_name="Optical Loss", startrow=2, startcol=2, index=False)  # send df to writer
    #            worksheet = OT_excel_file.sheets["Optical Loss"]  # pull worksheet object
    #            for idx, col in enumerate(combined_dataframe):  # loop through all columns
    #                series = combined_dataframe[col]
    #                max_len = max((
    #                    series.astype(str).map(len).max(),  # len of largest item
    #                    len(str(series.name))  # len of column name/header
    #                    )) + 8  # adding a little extra space
    #                worksheet.set_column(idx, idx, max_len)  # set column width



            # worksheet.set_column(Chn_col_list,Chn_col_list, None, bold)
            # worksheet.set_column(Slab_col_list,Slab_col_list, None, bold)
            # worksheet.set_column(SiN_col_list,SiN_col_list, None, bold)


            # Center_alignment_layer3.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
            # Center_alignment_layer3.api.VerticalAlignment = xw.constants.HAlign.xlHAlignCenter
            
                writer.close()
                self.button3.config(state=DISABLED)
                self.button2.config(state=DISABLED)

                # print("All Done!")
                os.startfile(OT_Save_Location+".xlsx")
            App.thread_inserttext(app, text="OT Results Saved!")
            App.thread_inserttext(app, text="--------------------------------------------------------------------------------------------------------------")
        except:
            App.thread_inserttext(app, text="...\n There is an error while extracting OT Data =(\nPossible issues:\n")

class Thk_Frame(ttk.Frame):
    # global Bottom_Widgets
    # global Bottom
    def __init__(self, container):
        super().__init__(container)
        self.columnconfigure(0, weight=1)
        self.columnconfigure(1, weight=3)
       
        self.create_Thk_widgets()
       
    def create_Thk_widgets(self):
        self.label = Label(self, text="This is for Thickness Test, please select the FILES.", background='#CCCCFF', width=80)
        self.label.grid(row=0, column=0, columnspan=3, sticky=E+W, padx=5, pady=5)      
       
        self.button1 = Button(self,
                              text="Open Thk Files (Without Saving):",
                              command=self.thread_openthkfile,
                              font=("Comic Sans",11),
                              relief=GROOVE,
                              state=ACTIVE)
        self.button1.grid(row=1, column=0, sticky=W, padx=1, pady=1)

    def thread_openthkfile(self):
        return Thread(daemon=True, target=self.OpenThkFile).start()
       
    def OpenThkFile(self):
        try:
            App.thread_inserttext(app, text="Thickness Test Selected!\nLow GOF (0.94) values shown in red font in Wafer Map.\n5 Points values shown in bold in Wafer Map.")
            desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
            Thk_filelist = []
            def Average(lst):
                return mean(lst)
        
        # Thk_temp_file = tempfile.NamedTemporaryFile(suffix='.xlsx')
        # print(Thk_temp_file)
            global Thk_writer
            global Thk_openfile
            
            # self.scrolled_text.insert(tk.END,"Hi")

            def Low_GOF_Checker(lst):
                if any(item < 0.94 for item in lst):
                    return "Yes"
                else:
                    return "No"
                        
        #ask for user input, return file names in a tuple
            Thk_openfile = filedialog.askopenfilenames(filetypes = [('CSV Files','*.csv'), ('Excel', '*.xlsx')], initialdir=r"\\172.16.0.99\CleanRoom\00. Thickness Measurement Data For Lot Owner Review")
            # Thk_basename_short = []
            for file in Thk_openfile:
                Thk_basename = Path(file).stem
                Thk_basename_split = Thk_basename.split("_")
                Thk_basename_short = Thk_basename_split[0]+"_"+Thk_basename_split[1]
                # App.thread_inserttext(app, text="File(s) selected: "+Thk_basename)
                Thk_filelist.append(Thk_basename)
                Thk_filelist_xlsx = [desktop + "/" + x + "_YZ_Macro.xlsx" for x in Thk_filelist]
            
            print(Thk_basename_split)
            print(type(Thk_basename_split))
            print(Thk_basename_short)
            print(type(Thk_basename_short))
            print(Thk_filelist_xlsx)
                
            for file in Thk_openfile:
                App.thread_inserttext(app, text="--------------------------------------------------------------------------------------------------------------")
                App.thread_inserttext(app, text="Data Extraction will begin for: "+file+"\n")
                count = 0 
                Thk_df = pd.read_csv(file, index_col=False)
                Max_no_of_slots = Thk_df["Cassette slot"].nunique()
                No_of_slots = pd.unique(Thk_df["Cassette slot"])

                print(f"Max no. of slots: {Max_no_of_slots}")
                App.thread_inserttext(app, text="Number of Slots: "+ str(No_of_slots))
                # print(f"List of slots: {No_of_slots}")

                Slot_df = Thk_df["Cassette slot"].value_counts()
                Slot_df_dict = Slot_df.to_dict()
                # print(Slot_df_dict)

                No_of_layers_float = Thk_df["Number of layers"].mean()
                No_of_layers = int(No_of_layers_float)
                App.thread_inserttext(app, text="Number of Layers: "+ str(No_of_layers))
                # print(f"No. of layers: {No_of_layers}")

                Max_site = Thk_df["Site"].max()
                App.thread_inserttext(app, text="Number of points measured per slot: "+ str(Max_site)+" Points")
                # print("This is max site:")
                # print(Max_site)

                Max_layers = Thk_df["Number of layers"].max()
                print(Max_layers)

                find_site = Thk_df.columns.get_loc("Site")



                Layer_thk = []
                Combined_thk = []
                layer_count= 1

                for i in range(No_of_layers):
                    for count,value in enumerate(Thk_df.iloc[:,find_site+layer_count]):
                        Layer_thk.append(value)
                    Combined_thk.append(Layer_thk)
                    Layer_thk=[]
                    layer_count+=1

                # print(f"This is combined thk of all slots: {Combined_thk}")            

            #Create empty list based on number of layers tested
                Layer_empty_list = [ [] for i in range(No_of_layers)]
                Layer_empty_list_for_pivot = [ [] for j in range(No_of_layers)]
                temp_slot_layer_pivot = []
                temp_slot_layer = []
                layer_count = 0

            #To iterate over EACH thk layers 1,2,3,...
                count_combined_thk = 0
                for i in Combined_thk:
                    length_combined_thk = len(Combined_thk[count_combined_thk])

                #In EACH layer, split it into EACH slot
                    for j in range(0, length_combined_thk, list(Slot_df_dict.values())[0]):
                        temp_slot_layer = i[j:j+list(Slot_df_dict.values())[0]]
                        temp_slot_layer_pivot = i[j:j+list(Slot_df_dict.values())[0]]
                        Layer_empty_list_for_pivot[layer_count].append(temp_slot_layer_pivot)
                        Layer_empty_list[layer_count].append(temp_slot_layer)
                        temp_slot_layer = []
                    layer_count+=1
                    count_combined_thk+=1
                #print(Layer_empty_list)
                    #Final result is a list per SLOTS per LAYERS

            #For GOF next
                GOF_list = []
                GOF_temp = []
                GOF_by_slot = []
                GOF_by_slot_avg = []
            #Get max no. of rows of GOF data + add them into a list
            #First check if Goodness-of-Fit column is present, because sometimes tool data abnormal, column name may disappear.
            #In that case, we check if it is 3 columns left from the end, which is usually the spot. But file format may have issues. So check if all values are < 1, which is typical of GOF values.
            #Otherwise, it will most likely be 2 columns left from the end instead. Thus the if-else statements to determine the location of GOF column.
            #Also, within each if-else statement, values with low GOF PER layer is sorted out separately, to add font red color in pivot table (wafer map) at the end
                Low_GOF_Thk_list = []

                if "Goodness-of-Fit" in Thk_df.columns:
                    GOF_rows = len(Thk_df["Goodness-of-Fit"].index)
                    for (columnName, columnData) in Thk_df["Goodness-of-Fit"].iteritems():
                        GOF_list.append(columnData)

                    Low_GOF_Thk_df = Thk_df[Thk_df["Goodness-of-Fit"]<0.94]
                    find_site_low_GOF = Low_GOF_Thk_df.columns.get_loc("Site")
                    for i in range(0, Max_layers):
                        Low_GOF_Thk_list.append(Low_GOF_Thk_df.iloc[:,find_site_low_GOF+1+i].tolist())
                    print(Low_GOF_Thk_list)

                elif (Thk_df.iloc[:,-3] < 1).all():
                    GOF_rows = len(Thk_df.iloc[:,-3])
                    for (columnName, columnData) in Thk_df.iloc[:,-3].iteritems():
                        GOF_list.append(columnData)

                    Low_GOF_Thk_df = Thk_df[Thk_df.iloc[:,-3]<0.94]
                    find_site_low_GOF = Low_GOF_Thk_df.columns.get_loc("Site")
                    for i in range(0, Max_layers):
                        Low_GOF_Thk_list.append(Low_GOF_Thk_df.iloc[:,find_site_low_GOF+1+i].tolist())
                    # print(Low_GOF_Thk_list)

                else:
                    GOF_rows = len(Thk_df.iloc[:,-2])
                    for (columnName, columnData) in Thk_df.iloc[:,-2].iteritems():
                        GOF_list.append(columnData)   
                    Low_GOF_Thk_df = Thk_df[Thk_df.iloc[:,-2]<0.94]
                    find_site_low_GOF = Low_GOF_Thk_df.columns.get_loc("Site")
                    for i in range(0, Max_layers):
                        Low_GOF_Thk_list.append(Low_GOF_Thk_df.iloc[:,find_site_low_GOF+1+i].tolist())         
                    # print(Low_GOF_Thk_list)               

                # No_low_GOF_filler = ["NoLowGOF"]*Max_layers
                # No_low_GOF_black_font = ["#000000"]*Max_layers

                # if len(Low_GOF_Thk_list) == 0:
                #     Low_GOF_red_font_list = No_low_GOF_black_font
                #     Low_GOF_Thk_list = No_low_GOF_filler
                # else:
                Low_GOF_red_font_list = ["#FF0000"]*len(Low_GOF_Thk_list[0])

                #In the list of GOF data, put them into list by each slot
                for i in range(0, GOF_rows, list(Slot_df_dict.values())[0]):
                    GOF_temp = GOF_list[i:i+list(Slot_df_dict.values())[0]]
                    GOF_by_slot.append(GOF_temp)
                    GOF_temp = []
                print(GOF_by_slot)

                for i in GOF_by_slot:
                    Avg_GOF = mean(i)
                    GOF_by_slot_avg.append(Avg_GOF)
                    Avg_GOF = 0
                # #print(GOF_by_slot_avg)
                
            
                GOF_df = pd.DataFrame(GOF_by_slot_avg, columns=['Avg GOF'])
                ##print(GOF_df)

            
            
            #Put GOF at the front + Thk at the back, so that it starts GOF as 0, and Thk counts from 1 onwards, no need to format
                for i in range(0, len(No_of_slots)):
                    for j in range(0, len(Layer_empty_list)):
                        Layer_empty_list[j][i].insert(0, GOF_by_slot_avg[i])
                ##print(f"This is new layer_empty_list: {Layer_empty_list}") 
            # [0 0] [1 0] [2 0]
            # [0 1] [1 1] [2 1]
            # [0 2] [1 2] [2 2]

            #Create list of names to use for naming pd.DataFrame variables based on no. of layers
                Temp_layer_create_count = 1
                Temp_layer_create_df = []
                for i in range(0, No_of_layers):
                    string = "Layer_df_"+str(Temp_layer_create_count)
                    Temp_layer_create_df.append(string)
                    Temp_layer_create_count+=1
                ##print(f"This is Temp_layer_create_df: {Temp_layer_create_df}")

                #Create pd.DataFrame based on no. of layers
                Temp_layer_count = 0
                for i in Layer_empty_list:
                    Temp_layer_create_df[Temp_layer_count] = pd.DataFrame(Layer_empty_list[Temp_layer_count], index=None)

                    ##print(f"This is Temp_layer_create_df[Temp_layer_count]: {Temp_layer_create_df[Temp_layer_count]}")
                    Temp_layer_count+=1

                wb_wings = xw.Book()

                ws_wings1 = wb_wings.sheets[0]
                ws_wings1.name = "Layer 1"
                ws_wings2 = wb_wings.sheets.add("Layer 2", after="Layer 1")
                ws_wings3 = wb_wings.sheets.add("Layer 3", after="Layer 2")
                ws_wings4 = wb_wings.sheets.add(Thk_basename_short, after="Layer 3")

                ws_wings1.activate()
                ws_wings4["A1"].options(index=False, expand='table').value = Thk_df

    #___________________________________________________________________________________________________________________
    # Start of Layer1 DF + Excel

                #Final dataframe, per layer
                #This is made for ALL thk points. 
                column_names = ["Recipe", "Lot ID", "Slot No."]
                combined_thk_all_pts_df_layer1 = pd.DataFrame(columns=column_names, index=None)
                combined_thk_all_pts_layer1 = pd.concat([combined_thk_all_pts_df_layer1, Temp_layer_create_df[0]])
                combined_thk_all_pts_layer1.rename(columns={combined_thk_all_pts_layer1.columns[3]:'Avg GOF'}, inplace=True)
                combined_thk_all_pts_layer1.loc[0,"Recipe"] = Thk_df.loc[0,"Jobfile Name"]
                combined_thk_all_pts_layer1.loc[:,"Lot ID"] = Thk_df.loc[:,"Lot ID"]
                for i in range(len(No_of_slots)):
                    combined_thk_all_pts_layer1.loc[i,"Slot No."] = No_of_slots[i]
                Avg_thk_list_layer1 = []
                for i in range(len(No_of_slots)):
                    Avg_thk_layer1 = combined_thk_all_pts_layer1.iloc[i,4:].mean(skipna=True)
                    Avg_thk_list_layer1.append(Avg_thk_layer1)
                combined_thk_all_pts_layer1 = combined_thk_all_pts_layer1.assign(Avg_Thk_All_Pts=Avg_thk_list_layer1)
                ##print("This is combined_thk_all_pts_layer1: ")
                ##print(combined_thk_all_pts_layer1)
                combined_thk_all_pts_layer1["Avg GOF"] = combined_thk_all_pts_layer1["Avg GOF"].round(decimals=3)
                Thk_all_pts_1decimal_layer1 = combined_thk_all_pts_layer1.columns.get_loc("Avg GOF")
                combined_thk_all_pts_layer1.iloc[:,Thk_all_pts_1decimal_layer1+1:] = combined_thk_all_pts_layer1.iloc[:,Thk_all_pts_1decimal_layer1+1:].round(decimals=1)
                combined_thk_all_pts_layer1["Any Low GOF (< 0.94)?"] = ""
                Low_GOF_list = []
                Low_gof_count = 0

                for i in range(0, len(No_of_slots)):
                    Low_GOF_Boolean = Low_GOF_Checker(GOF_by_slot[i])
                    Low_GOF_list.append(Low_GOF_Boolean)
                print("Low GOF List:")
                print(Low_GOF_list)
                combined_thk_all_pts_layer1.iloc[:,-1]=Low_GOF_list

                #-------------------------------------------------------------------------------------------------------------------------------------------
                # #write for 1-5 pt next

                Temp_thk_1to5_layer1 = []
                Temp_thk_1to5_j_layer1 = []
                Temp_layer_empty_list_1to5_layer1 = []
                Layer_empty_list_1to5_layer1 = []

                for i in Layer_empty_list:
                    for j in range(len(i)):
                        Temp_thk_1to5_j_layer1 = i[j][:6]
                        Temp_thk_1to5_layer1.append(Temp_thk_1to5_j_layer1)
                ##print(f"This is Temp_thk_1to5_layer1:\n {Temp_thk_1to5_layer1}")
                ##print(len(Temp_thk_1to5_layer1))
                ##print(len(Temp_thk_1to5_layer1))
                ##print(len(Layer_empty_list))
                for i in range(0, len(Temp_thk_1to5_layer1), int(len(Temp_thk_1to5_layer1)/len(Layer_empty_list))):
                    Temp_layer_empty_list_1to5_layer1 = Temp_thk_1to5_layer1[i:i+int(len(Temp_thk_1to5_layer1)/len(Layer_empty_list))]
                    Layer_empty_list_1to5_layer1.append(Temp_layer_empty_list_1to5_layer1)
                    
                ##print(f"This is Layer_empty_list_1to5_layer1:\n {Layer_empty_list_1to5_layer1}")
                #Example to reduce [1,2,3,4,5] to [1,2,3] for each list of lists. len(list1) = No. of layers, len(list1)[i] = No. of points in each slot in each layer
                # list1 = [[[1,2,3,4,5],[1,2,3,4,5],[1,2,3,4,5]],[[10,20,30,40,50],[10,20,30,40,50],[10,20,30,40,50]]]
                # list2 = []
                # list3 = []
                # list4 = []
                # for i in list1:
                #     for j in range(len(i)):
                #         list2.append(i[j][:4])
                #         ##print(f"This is list2: {list2}")
                # ##print(len(list2))
                # for i in range(0, len(list2), int(len(list2)/len(list1))):
                #     list3 = list2[i:i+int(len(list2)/len(list1))]
                #     ##print(f"This is list3: {list3}")
                #     list4.append(list3)
                # ##print(f"This is list4: {list4}")

                #Create list of names to use for naming pd.DataFrame variables based on no. of layers
                Temp_layer_create_count_1to5_layer1 = 1
                Temp_layer_create_1to5_df_layer1 = []
                for i in range(0, No_of_layers):
                    string_layer1 = "Layer_df_"+str(Temp_layer_create_count_1to5_layer1)+"_1to5"
                    Temp_layer_create_1to5_df_layer1.append(string_layer1)
                    Temp_layer_create_count_1to5_layer1+=1
                ##print("This is Temp_layer_create_1to5_df_layer1: ")
                ##print(Temp_layer_create_1to5_df_layer1)

                #Create pd.DataFrame based on no. of layers
                Temp_layer_count_1to5_layer1 = 0
                for i in Layer_empty_list_1to5_layer1:
                    Temp_layer_create_1to5_df_layer1[Temp_layer_count_1to5_layer1] = pd.DataFrame(Layer_empty_list_1to5_layer1[Temp_layer_count_1to5_layer1], index=None)

                    ##print("This is Temp_layer_create_1to5_df_layer1[Temp_layer_count]: ") 
                    ##print(Temp_layer_create_1to5_df_layer1[Temp_layer_count_1to5_layer1])
                    Temp_layer_count_1to5_layer1+=1

                #At this point, for loop to create sheets for each layer
                combined_thk_1to5_df_layer1 = pd.DataFrame(columns=column_names, index=None)
                combined_thk_1to5_layer1 = pd.concat([combined_thk_1to5_df_layer1, Temp_layer_create_1to5_df_layer1[0]])
                # combined_thk_1to5_layer1.style.set_properties(**{'text-align': 'center'})
                combined_thk_1to5_layer1.rename(columns={combined_thk_1to5_layer1.columns[3]:'Avg GOF'}, inplace=True)
                combined_thk_1to5_layer1.loc[0,"Recipe"] = Thk_df.loc[0,"Jobfile Name"]
                combined_thk_1to5_layer1.loc[:,"Lot ID"] = Thk_df.loc[:,"Lot ID"]
                for i in range(len(No_of_slots)):
                    combined_thk_1to5_layer1.loc[i,"Slot No."] = No_of_slots[i]
                Avg_thk_list_1to5_layer1 = []
                for i in range(len(No_of_slots)):
                    Avg_thk_1to5_layer1 = combined_thk_1to5_layer1.iloc[i,4:].mean(skipna=True)
                    Avg_thk_list_1to5_layer1.append(Avg_thk_1to5_layer1)
                combined_thk_1to5_layer1 = combined_thk_1to5_layer1.assign(Avg_Thk_5pts=Avg_thk_list_1to5_layer1)
                ##print("This is combined_thk_1to5_layer1: ")
                ##print(combined_thk_1to5_layer1)      
                combined_thk_1to5_layer1["Avg GOF"] = combined_thk_1to5_layer1["Avg GOF"].round(decimals=3)
                Thk_1to5_1decimal_layer1 = combined_thk_1to5_layer1.columns.get_loc("Avg GOF")
                combined_thk_1to5_layer1.iloc[:,Thk_1to5_1decimal_layer1+1:] = combined_thk_1to5_layer1.iloc[:,Thk_1to5_1decimal_layer1+1:].round(decimals=1)

                #-------------------------------------------------------------------------------------------------------------    
                #All the dataframe placement done for both: 1) All pts and 2) 5 Pts. Time to put in excel file using xlwings (Because it can create temp workbook)

                row_size_1to5 = int(combined_thk_1to5_layer1.shape[0])
                col_size_1to5 = int(combined_thk_1to5_layer1.shape[1])

                row_size_all_pts = int(combined_thk_all_pts_layer1.shape[0])
                col_size_all_pts = int(combined_thk_all_pts_layer1.shape[1])

                row_spacing_1to5_all_pts = int(combined_thk_1to5_layer1.shape[0]) + 2
                col_spacing_1to5_all_pts = int(combined_thk_1to5_layer1.shape[1])

                ws_wings1["B1"].value = "5 Points Thickness: "
                ws_wings1["B1"].font.bold = True
                ws_wings1.range(2,2).options(index=False, expand='table').value = combined_thk_1to5_layer1
                ws_wings1.range((3,2),(2+row_size_1to5,2)).merge()
                ws_wings1.range((2,2),(2+row_size_1to5,2+col_size_1to5-1)).api.Borders.Weight = 2
                header_1to5_layer1 = ws_wings1.range(2,2).expand('right')
                header_1to5_layer1.color = (209,252,237)
                header_1to5_layer1.font.bold = True
                ws_wings1.range(2+int(row_spacing_1to5_all_pts),2).value = "All Points Thickness: "
                ws_wings1.range(2+int(row_spacing_1to5_all_pts),2).font.bold = True
                ws_wings1.range(2+int(row_spacing_1to5_all_pts)+1,2).options(index=False).value = combined_thk_all_pts_layer1
                ws_wings1.range((2+int(row_spacing_1to5_all_pts)+1+1,2),(2+int(row_spacing_1to5_all_pts)+1+row_size_all_pts,2)).merge()
                ws_wings1.range(2+int(row_spacing_1to5_all_pts)+1+int(row_size_all_pts)+2,2).value = "Wafer Points on Map:"
                ws_wings1.range(2+int(row_spacing_1to5_all_pts)+1+int(row_size_all_pts)+2,2).font.bold = True
                Pivot_row1 = 2+int(row_spacing_1to5_all_pts)+1+int(row_size_all_pts)+2
                Pivot_col1 = 2

                header_all_pts_layer1 = ws_wings1.range(2+int(row_spacing_1to5_all_pts)+1,2).expand('right')
                header_all_pts_layer1.color = (209,252,237)
                header_all_pts_layer1.font.bold = True
                ws_wings1.range((2+int(row_spacing_1to5_all_pts)+1,2),(2+int(row_spacing_1to5_all_pts)+1+row_size_all_pts,2+col_size_all_pts-1)).api.Borders.Weight = 2

                #Pivot Table for Layer 1 Thk
                #Create list of names to use for naming PIVOT TABLE based on no. of slots for LAYER 1
                Pivot_slot_layer1 = []
                for i in range(0, Max_no_of_slots):
                    string_layer1 = "PivotTable_Slot"+str(No_of_slots[i])
                    Pivot_slot_layer1.append(string_layer1)
                #print(f"Pivot_slot_layer1: {Pivot_slot_layer1}")

                #Create list of names for pivot table DATAFRAME based on no. of slots for LAYER1, to iterate during pivot table creation
                Pivot_df_count_layer1 = 1
                Pivot_df_layer1 = []
                for i in range(0, Max_no_of_slots):
                    string_layer1 = "Layer_df_"+str(No_of_slots[i])
                    Pivot_df_layer1.append(string_layer1)
                    Pivot_df_count_layer1+=1

                Max_site_count_layer1 = 0 
                #Create pivot df with Thk_df "X Pos", "Y Pos", "Layer 1 Thickness"
                for i in range(0, len(Pivot_df_layer1)):
                    Pivot_df_layer1[i] = Thk_df.loc[Max_site_count_layer1*Max_site:Max_site-1+Max_site_count_layer1*Max_site,["X Pos", "Y Pos", "Layer 1 Thickness"]]
                    #print(Pivot_df_layer1[i])
                    Max_site_count_layer1+=1

                pivot_space_count_layer1 = 1
                pivot_space_count2_layer1 = 0
                for i in range(0, Max_no_of_slots):
                    Pivot_slot_layer1[i] = pd.pivot_table(Pivot_df_layer1[i], values="Layer 1 Thickness", index="Y Pos", columns="X Pos", aggfunc="mean")
                    row_size_pivot = Pivot_slot_layer1[i].shape[0]
                    col_size_pivot = Pivot_slot_layer1[i].shape[1]
                    #print(f"row_size_pivot = {row_size_pivot}")
                    ws_wings1.range(Pivot_row1+row_size_pivot*i+pivot_space_count_layer1+pivot_space_count2_layer1,Pivot_col1).value = Pivot_slot_layer1[i]
                    temp_pivot_header_col_layer1 = ws_wings1.range(Pivot_row1+row_size_pivot*i+pivot_space_count_layer1+pivot_space_count2_layer1,Pivot_col1).expand('right')
                    temp_pivot_header_col_layer1.color = (209,252,237)
                    temp_pivot_header_col_layer1.api.Borders.Weight=2
                    temp_pivot_header_row_layer1 = ws_wings1.range(Pivot_row1+row_size_pivot*i+pivot_space_count_layer1+pivot_space_count2_layer1,Pivot_col1).expand('down')
                    temp_pivot_header_row_layer1.color = (209,252,237)
                    temp_pivot_header_row_layer1.api.Borders.Weight=2

                    
                    Pivot_layer1_thk_list = Pivot_df_layer1[i]["Layer 1 Thickness"].tolist()
                    Pivot_layer1_thk_list_dcopy = copy.deepcopy(Pivot_layer1_thk_list)
                    Pivot_layer1_thk_list.sort()

                    Pivot_5pts_list_layer1 = []
                    Pivot_bold_list_layer1 = [True]*5
                    # print(f"Pivot_bold_list_layer1: {Pivot_bold_list_layer1}")

                    Pivot_5pts_dict_layer1 = {}

                    for l in range(0, len(Pivot_layer1_thk_list_dcopy), Max_site):
                        Pivot_5pts_list_layer1.append(Pivot_layer1_thk_list_dcopy[l:l+5])
                    # print(f"This is Pivot_5pts_list: {Pivot_5pts_list_layer1}")                    

                    Pivot_5pts_dict_layer1 = dict(zip(Pivot_5pts_list_layer1[0],Pivot_bold_list_layer1))
                    print(Pivot_5pts_dict_layer1)

                    #Create color scale for pivot table values
                    Red_div_layer1 = (255-0)/len(Pivot_layer1_thk_list)
                    Red_div_rounded_layer1 = np.floor(Red_div_layer1)
                    ##print(Red_div_rounded_layer1)

                    Blue_div_layer1 = (255-179)/len(Pivot_layer1_thk_list)
                    Blue_div_rounded_layer1 = np.floor(Blue_div_layer1)
                    ##print(Blue_div_rounded_layer1)

                    Pivot_color_scale_layer1 = []
                    for j in range(0,len(Pivot_layer1_thk_list)):
                        temp_pivot_tuple = (int(255-j*Red_div_rounded_layer1),int(255),int(255-j*Blue_div_rounded_layer1))
                        Pivot_color_scale_layer1.append(temp_pivot_tuple)
                    ##print(Pivot_color_scale_layer1)
                    
                    Pivot_color_scale_dict_layer1 = {}
                    Pivot_color_scale_dict_layer1 = dict(zip(Pivot_layer1_thk_list,Pivot_color_scale_layer1))
                    ##print(Pivot_color_scale_dict_layer1)

                    #Add red font to low GOF values in wafer map/pivot table
                    Low_GOF_pivot_red_font_dict_layer1 = {}
                    Low_GOF_pivot_red_font_dict_layer1 = dict(zip(Low_GOF_Thk_list[0],Low_GOF_red_font_list))


                    for k in ws_wings1.range((Pivot_row1+row_size_pivot*i+pivot_space_count_layer1+pivot_space_count2_layer1+1,Pivot_col1+1),(Pivot_row1+row_size_pivot*i+pivot_space_count_layer1+pivot_space_count2_layer1+row_size_pivot,Pivot_col1+col_size_pivot)):
                        if k.value is not None:
                            k.color = Pivot_color_scale_dict_layer1.get(k.value)
                            k.font.bold = Pivot_5pts_dict_layer1.get(k.value)
                            try:
                                k.font.color = Low_GOF_pivot_red_font_dict_layer1.get(k.value)
                            except:
                                pass
                        else:
                            continue


                    ws_wings1.range(Pivot_row1+row_size_pivot*i+pivot_space_count_layer1+pivot_space_count2_layer1,Pivot_col1).value = "Slot "+ str(No_of_slots[i])
                    ws_wings1.range(Pivot_row1+row_size_pivot*i+pivot_space_count_layer1+pivot_space_count2_layer1,Pivot_col1).font.bold = True
                    pivot_space_count2_layer1+=1
                    pivot_space_count_layer1+=1
                #End of pivot table + color scale function

                Center_alignment_layer1 = ws_wings1.range((1,1),(Pivot_row1+row_size_pivot*Max_no_of_slots+20,1+col_size_all_pts+10))
                Center_alignment_layer1.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
                Center_alignment_layer1.api.VerticalAlignment = xw.constants.HAlign.xlHAlignCenter
                ws_wings1.autofit(axis="columns")
                App.thread_inserttext(app, text="...Layer 1 Done!")
                print("No of layers: "+str(No_of_layers))

    #End of LAYER1 EXCEL
    #################################################################################################
    #Start of Layer2 DF + Excel
                if No_of_layers > 1:
                    #Final dataframe, per layer
                    #This is made for ALL thk points. 
                    column_names = ["Recipe", "Lot ID", "Slot No."]
                    combined_thk_all_pts_df_layer2 = pd.DataFrame(columns=column_names, index=None)
                    combined_thk_all_pts_layer2 = pd.concat([combined_thk_all_pts_df_layer2, Temp_layer_create_df[1]])
                    combined_thk_all_pts_layer2.rename(columns={combined_thk_all_pts_layer2.columns[3]:'Avg GOF'}, inplace=True)
                    combined_thk_all_pts_layer2.loc[0,"Recipe"] = Thk_df.loc[0,"Jobfile Name"]
                    combined_thk_all_pts_layer2.loc[:,"Lot ID"] = Thk_df.loc[:,"Lot ID"]
                    for i in range(len(No_of_slots)):
                        combined_thk_all_pts_layer2.loc[i,"Slot No."] = No_of_slots[i]
                    Avg_thk_list_layer2 = []
                    for i in range(len(No_of_slots)):
                        Avg_thk_layer2 = combined_thk_all_pts_layer2.iloc[i,4:].mean(skipna=True)
                        Avg_thk_list_layer2.append(Avg_thk_layer2)
                    combined_thk_all_pts_layer2 = combined_thk_all_pts_layer2.assign(Avg_Thk_All_Pts=Avg_thk_list_layer2)
                    ##print("This is combined_thk_all_pts_layer2: ")
                    ##print(combined_thk_all_pts_layer2)
                    combined_thk_all_pts_layer2["Avg GOF"] = combined_thk_all_pts_layer2["Avg GOF"].round(decimals=3)
                    Thk_all_pts_1decimal_layer2 = combined_thk_all_pts_layer2.columns.get_loc("Avg GOF")
                    combined_thk_all_pts_layer2.iloc[:,Thk_all_pts_1decimal_layer2+1:] = combined_thk_all_pts_layer2.iloc[:,Thk_all_pts_1decimal_layer2+1:].round(decimals=1)
                    combined_thk_all_pts_layer2["Any Low GOF (< 0.94)?"] = ""
                    combined_thk_all_pts_layer2.iloc[:,-1]=Low_GOF_list

                    #-------------------------------------------------------------------------------------------------------------------------------------------
                    # #write for 1-5 pt next

                    Temp_thk_1to5_layer2 = []
                    Temp_thk_1to5_j_layer2 = []
                    Temp_layer_empty_list_1to5_layer2 = []
                    Layer_empty_list_1to5_layer2 = []

                    for i in Layer_empty_list:
                        for j in range(len(i)):
                            Temp_thk_1to5_j_layer2 = i[j][:6]
                            Temp_thk_1to5_layer2.append(Temp_thk_1to5_j_layer2)
                    for i in range(0, len(Temp_thk_1to5_layer2), int(len(Temp_thk_1to5_layer2)/len(Layer_empty_list))):
                        Temp_layer_empty_list_1to5_layer2 = Temp_thk_1to5_layer2[i:i+int(len(Temp_thk_1to5_layer2)/len(Layer_empty_list))]
                        Layer_empty_list_1to5_layer2.append(Temp_layer_empty_list_1to5_layer2)

                    #Create list of names to use for naming pd.DataFrame variables based on no. of layers
                    Temp_layer_create_count_1to5_layer2 = 1
                    Temp_layer_create_1to5_df_layer2 = []
                    for i in range(0, No_of_layers):
                        string_layer2 = "Layer_df_"+str(Temp_layer_create_count_1to5_layer2)+"_1to5"
                        Temp_layer_create_1to5_df_layer2.append(string_layer2)
                        Temp_layer_create_count_1to5_layer2+=1
                    ##print("This is Temp_layer_create_1to5_df_layer2: ")
                    ##print(Temp_layer_create_1to5_df_layer2)

                    #Create pd.DataFrame based on no. of layers
                    Temp_layer_count_1to5_layer2 = 0
                    for i in Layer_empty_list_1to5_layer2:
                        Temp_layer_create_1to5_df_layer2[Temp_layer_count_1to5_layer2] = pd.DataFrame(Layer_empty_list_1to5_layer2[Temp_layer_count_1to5_layer2], index=None)

                        ##print("This is Temp_layer_create_1to5_df_layer2[Temp_layer_count]: ") 
                        ##print(Temp_layer_create_1to5_df_layer2[Temp_layer_count_1to5_layer2])
                        Temp_layer_count_1to5_layer2+=1

                    #At this point, for loop to create sheets for each layer
                    combined_thk_1to5_df_layer2 = pd.DataFrame(columns=column_names, index=None)
                    combined_thk_1to5_layer2 = pd.concat([combined_thk_1to5_df_layer2, Temp_layer_create_1to5_df_layer2[1]])
                    combined_thk_1to5_layer2.rename(columns={combined_thk_1to5_layer2.columns[3]:'Avg GOF'}, inplace=True)
                    combined_thk_1to5_layer2.loc[0,"Recipe"] = Thk_df.loc[0,"Jobfile Name"]
                    combined_thk_1to5_layer2.loc[:,"Lot ID"] = Thk_df.loc[:,"Lot ID"]
                    for i in range(len(No_of_slots)):
                        combined_thk_1to5_layer2.loc[i,"Slot No."] = No_of_slots[i]
                    Avg_thk_list_1to5_layer2 = []
                    for i in range(len(No_of_slots)):
                        Avg_thk_1to5_layer2 = combined_thk_1to5_layer2.iloc[i,4:].mean(skipna=True)
                        Avg_thk_list_1to5_layer2.append(Avg_thk_1to5_layer2)
                    combined_thk_1to5_layer2 = combined_thk_1to5_layer2.assign(Avg_Thk_5pts=Avg_thk_list_1to5_layer2)
                    ##print("This is combined_thk_1to5_layer2: ")
                    ##print(combined_thk_1to5_layer2)      
                    combined_thk_1to5_layer2["Avg GOF"] = combined_thk_1to5_layer2["Avg GOF"].round(decimals=3)
                    # combined_thk_1to5_layer2["Avg_thk_5pts"] = combined_thk_1to5_layer2["Avg_thk_5pts"].round(decimals=1)
                    Thk_1to5_1decimal_layer2 = combined_thk_1to5_layer2.columns.get_loc("Avg GOF")
                    combined_thk_1to5_layer2.iloc[:,Thk_1to5_1decimal_layer2+1:] = combined_thk_1to5_layer2.iloc[:,Thk_1to5_1decimal_layer2+1:].round(decimals=1)

                    #-------------------------------------------------------------------------------------------------------------    
                    #All the dataframe placement done for both: 1) All pts and 2) 5 Pts. Time to put in excel file using xlwings (Because it can create temp workbook)

                    row_size_1to5 = int(combined_thk_1to5_layer2.shape[0])
                    col_size_1to5 = int(combined_thk_1to5_layer2.shape[1])

                    row_size_all_pts = int(combined_thk_all_pts_layer2.shape[0])
                    col_size_all_pts = int(combined_thk_all_pts_layer2.shape[1])

                    row_spacing_1to5_all_pts = int(combined_thk_1to5_layer2.shape[0]) + 2
                    col_spacing_1to5_all_pts = int(combined_thk_1to5_layer2.shape[1])

                    ws_wings2["B1"].value = "5 Points Thickness: "
                    ws_wings2["B1"].font.bold = True
                    ws_wings2.range(2,2).options(index=False, expand='table').value = combined_thk_1to5_layer2
                    ws_wings2.range((3,2),(2+row_size_1to5,2)).merge()
                    ws_wings2.range((2,2),(2+row_size_1to5,2+col_size_1to5-1)).api.Borders.Weight = 2
                    header_1to5_layer2 = ws_wings2.range(2,2).expand('right')
                    header_1to5_layer2.color = (209,252,237)
                    header_1to5_layer2.font.bold = True
                    ws_wings2.range(2+int(row_spacing_1to5_all_pts),2).value = "All Points Thickness: "
                    ws_wings2.range(2+int(row_spacing_1to5_all_pts),2).font.bold = True
                    ws_wings2.range(2+int(row_spacing_1to5_all_pts)+1,2).options(index=False).value = combined_thk_all_pts_layer2
                    ws_wings2.range((2+int(row_spacing_1to5_all_pts)+1+1,2),(2+int(row_spacing_1to5_all_pts)+1+row_size_all_pts,2)).merge()
                    ws_wings2.range(2+int(row_spacing_1to5_all_pts)+1+int(row_size_all_pts)+2,2).value = "Wafer Points on Map:"
                    ws_wings2.range(2+int(row_spacing_1to5_all_pts)+1+int(row_size_all_pts)+2,2).font.bold = True
                    #Pivot_row1 = 2+int(row_spacing_1to5_all_pts)+1+int(row_size_all_pts)+2
                    #Pivot_col1 = 2                

                    header_all_pts_layer2 = ws_wings2.range(2+int(row_spacing_1to5_all_pts)+1,2).expand('right')
                    header_all_pts_layer2.color = (209,252,237)
                    header_all_pts_layer2.font.bold = True
                    ws_wings2.range((2+int(row_spacing_1to5_all_pts)+1,2),(2+int(row_spacing_1to5_all_pts)+1+row_size_all_pts,2+col_size_all_pts-1)).api.Borders.Weight = 2

                    #Pivot Table for Layer 2 Thk
                    #Create list of names to use for naming PIVOT TABLE based on no. of slots for LAYER 1
                    Pivot_slot_layer2 = []
                    for i in range(0, Max_no_of_slots):
                        string_layer2 = "PivotTable_Slot"+str(No_of_slots[i])
                        Pivot_slot_layer2.append(string_layer2)
                    #print(f"Pivot_slot_layer2: {Pivot_slot_layer2}")


                    #Create list of names for pivot table DATAFRAME based on no. of slots for LAYER1, to iterate during pivot table creation
                    Pivot_df_count_layer2 = 1
                    Pivot_df_layer2 = []
                    for i in range(0, Max_no_of_slots):
                        string_layer2 = "Layer_df_"+str(No_of_slots[i])
                        Pivot_df_layer2.append(string_layer2)
                        Pivot_df_count_layer2+=1

                    Max_site_count_layer2 = 0 
                    #Create pivot df with Thk_df "X Pos", "Y Pos", "Layer 2 Thickness"
                    for i in range(0, len(Pivot_df_layer2)):
                        Pivot_df_layer2[i] = Thk_df.loc[Max_site_count_layer2*Max_site:Max_site-1+Max_site_count_layer2*Max_site,["X Pos", "Y Pos", "Layer 2 Thickness"]]
                        #print(Pivot_df_layer2[i])
                        Max_site_count_layer2+=1

                    pivot_space_count_layer2 = 1
                    pivot_space_count2_layer2 = 0
                    for i in range(0, Max_no_of_slots):
                        Pivot_slot_layer2[i] = pd.pivot_table(Pivot_df_layer2[i], values="Layer 2 Thickness", index="Y Pos", columns="X Pos", aggfunc="mean")
                        row_size_pivot = Pivot_slot_layer2[i].shape[0]
                        col_size_pivot = Pivot_slot_layer2[i].shape[1]
                        #print(f"row_size_pivot = {row_size_pivot}")
                        ws_wings2.range(Pivot_row1+row_size_pivot*i+pivot_space_count_layer2+pivot_space_count2_layer2,Pivot_col1).value = Pivot_slot_layer2[i]
                        temp_pivot_header_col_layer2 = ws_wings2.range(Pivot_row1+row_size_pivot*i+pivot_space_count_layer2+pivot_space_count2_layer2,Pivot_col1).expand('right')
                        temp_pivot_header_col_layer2.color = (209,252,237)
                        temp_pivot_header_col_layer2.api.Borders.Weight=2
                        temp_pivot_header_row_layer2 = ws_wings2.range(Pivot_row1+row_size_pivot*i+pivot_space_count_layer2+pivot_space_count2_layer2,Pivot_col1).expand('down')
                        temp_pivot_header_row_layer2.color = (209,252,237)
                        temp_pivot_header_row_layer2.api.Borders.Weight=2

                        #Create color scale for pivot table values
                        Pivot_layer2_thk_list = Pivot_df_layer2[i]["Layer 2 Thickness"].tolist()
                        Pivot_layer2_thk_list_dcopy = copy.deepcopy(Pivot_layer2_thk_list)
                        Pivot_layer2_thk_list.sort()
                        ##print("list..")
                        ##print(Pivot_layer2_thk_list)
                        ##print(len(Pivot_layer2_thk_list))

                        Pivot_5pts_list_layer2 = []
                        Pivot_bold_list_layer2 = [True]*5
                        Pivot_5pts_dict_layer2 = {}

                        for l in range(0, len(Pivot_layer2_thk_list_dcopy), Max_site):
                            Pivot_5pts_list_layer2.append(Pivot_layer2_thk_list_dcopy[l:l+5])
                        
                        Pivot_5pts_dict_layer2 = dict(zip(Pivot_5pts_list_layer2[0],Pivot_bold_list_layer2))

                        Red_div_layer2 = (255-0)/len(Pivot_layer2_thk_list)
                        Red_div_rounded_layer2 = np.floor(Red_div_layer2)
                        ##print(Red_div_rounded_layer2)

                        Blue_div_layer2 = (255-179)/len(Pivot_layer2_thk_list)
                        Blue_div_rounded_layer2 = np.floor(Blue_div_layer2)
                        ##print(Blue_div_rounded_layer2)

                        Pivot_color_scale_layer2 = []
                        for j in range(0,len(Pivot_layer2_thk_list)):
                            temp_pivot_tuple = (int(255-j*Red_div_rounded_layer2),int(255),int(255-j*Blue_div_rounded_layer2))
                            Pivot_color_scale_layer2.append(temp_pivot_tuple)
                        ##print(Pivot_color_scale_layer2)
                        
                        Pivot_color_scale_dict_layer2 = {}
                        Pivot_color_scale_dict_layer2 = dict(zip(Pivot_layer2_thk_list,Pivot_color_scale_layer2))
                        ##print(Pivot_color_scale_dict_layer2)

                        #Add red font to low GOF values in wafer map/pivot table
                        Low_GOF_pivot_red_font_dict_layer2 = {}
                        Low_GOF_pivot_red_font_dict_layer2 = dict(zip(Low_GOF_Thk_list[1],Low_GOF_red_font_list))

                        for k in ws_wings2.range((Pivot_row1+row_size_pivot*i+pivot_space_count_layer2+pivot_space_count2_layer2+1,Pivot_col1+1),(Pivot_row1+row_size_pivot*i+pivot_space_count_layer2+pivot_space_count2_layer2+row_size_pivot,Pivot_col1+col_size_pivot)):
                            if k.value is not None:
                                k.color = Pivot_color_scale_dict_layer2.get(k.value)
                                k.font.bold = Pivot_5pts_dict_layer2.get(k.value)
                                try:
                                    k.font.color = Low_GOF_pivot_red_font_dict_layer2.get(k.value)
                                except:
                                    pass
                            else:
                                continue

                        ws_wings2.range(Pivot_row1+row_size_pivot*i+pivot_space_count_layer2+pivot_space_count2_layer2,Pivot_col1).value = "Slot "+ str(No_of_slots[i])
                        ws_wings2.range(Pivot_row1+row_size_pivot*i+pivot_space_count_layer2+pivot_space_count2_layer2,Pivot_col1).font.bold = True
                        pivot_space_count2_layer2+=1
                        pivot_space_count_layer2+=1
                    #End of pivot table + color scale function

                    Center_alignment_layer2 = ws_wings2.range((1,1),(Pivot_row1+row_size_pivot*Max_no_of_slots+20,1+col_size_all_pts+10))
                    Center_alignment_layer2.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
                    Center_alignment_layer2.api.VerticalAlignment = xw.constants.HAlign.xlHAlignCenter
                    ws_wings2.autofit(axis="columns")
                    App.thread_inserttext(app, text="...Layer 2 Done!")

    #End of Layer 2 Excel
    ###############################################################################
    #Start of Layer 3 DF + Excel
                    if No_of_layers > 2:
                        #Final dataframe, per layer
                        #This is made for ALL thk points. 
                        column_names = ["Recipe", "Lot ID", "Slot No."]
                        combined_thk_all_pts_df_layer3 = pd.DataFrame(columns=column_names, index=None)
                        combined_thk_all_pts_layer3 = pd.concat([combined_thk_all_pts_df_layer3, Temp_layer_create_df[2]])
                        combined_thk_all_pts_layer3.rename(columns={combined_thk_all_pts_layer3.columns[3]:'Avg GOF'}, inplace=True)
                        combined_thk_all_pts_layer3.loc[0,"Recipe"] = Thk_df.loc[0,"Jobfile Name"]
                        combined_thk_all_pts_layer3.loc[:,"Lot ID"] = Thk_df.loc[:,"Lot ID"]
                        for i in range(len(No_of_slots)):
                            combined_thk_all_pts_layer3.loc[i,"Slot No."] = No_of_slots[i]
                        Avg_thk_list_layer3 = []
                        for i in range(len(No_of_slots)):
                            Avg_thk_layer3 = combined_thk_all_pts_layer3.iloc[i,4:].mean(skipna=True)
                            Avg_thk_list_layer3.append(Avg_thk_layer3)
                        combined_thk_all_pts_layer3 = combined_thk_all_pts_layer3.assign(Avg_Thk_All_Pts=Avg_thk_list_layer3)
                        ##print("This is combined_thk_all_pts_layer3: ")
                        ##print(combined_thk_all_pts_layer3)
                        combined_thk_all_pts_layer3["Avg GOF"] = combined_thk_all_pts_layer3["Avg GOF"].round(decimals=3)
                        Thk_all_pts_1decimal_layer3 = combined_thk_all_pts_layer3.columns.get_loc("Avg GOF")
                        combined_thk_all_pts_layer3.iloc[:,Thk_all_pts_1decimal_layer3+1:] = combined_thk_all_pts_layer3.iloc[:,Thk_all_pts_1decimal_layer3+1:].round(decimals=1)
                        combined_thk_all_pts_layer3["Any Low GOF (< 0.94)?"] = ""
                        combined_thk_all_pts_layer3.iloc[:,-1]=Low_GOF_list
                        #-------------------------------------------------------------------------------------------------------------------------------------------
                        # #write for 1-5 pt next

                        Temp_thk_1to5_layer3 = []
                        Temp_thk_1to5_j_layer3 = []
                        Temp_layer_empty_list_1to5_layer3 = []
                        Layer_empty_list_1to5_layer3 = []

                        for i in Layer_empty_list:
                            for j in range(len(i)):
                                Temp_thk_1to5_j_layer3 = i[j][:6]
                                Temp_thk_1to5_layer3.append(Temp_thk_1to5_j_layer3)

                        for i in range(0, len(Temp_thk_1to5_layer3), int(len(Temp_thk_1to5_layer3)/len(Layer_empty_list))):
                            Temp_layer_empty_list_1to5_layer3 = Temp_thk_1to5_layer3[i:i+int(len(Temp_thk_1to5_layer3)/len(Layer_empty_list))]
                            Layer_empty_list_1to5_layer3.append(Temp_layer_empty_list_1to5_layer3)

                        #Create list of names to use for naming pd.DataFrame variables based on no. of layers
                        Temp_layer_create_count_1to5_layer3 = 1
                        Temp_layer_create_1to5_df_layer3 = []
                        for i in range(0, No_of_layers):
                            string_layer3 = "Layer_df_"+str(Temp_layer_create_count_1to5_layer3)+"_1to5"
                            Temp_layer_create_1to5_df_layer3.append(string_layer3)
                            Temp_layer_create_count_1to5_layer3+=1
                        ##print("This is Temp_layer_create_1to5_df_layer3: ")
                        ##print(Temp_layer_create_1to5_df_layer3)

                    #Create pd.DataFrame based on no. of layers
                        Temp_layer_count_1to5_layer3 = 0
                        for i in Layer_empty_list_1to5_layer3:
                            Temp_layer_create_1to5_df_layer3[Temp_layer_count_1to5_layer3] = pd.DataFrame(Layer_empty_list_1to5_layer3[Temp_layer_count_1to5_layer3], index=None)

                            Temp_layer_count_1to5_layer3+=1

                        #At this point, for loop to create sheets for each layer
                        combined_thk_1to5_df_layer3 = pd.DataFrame(columns=column_names, index=None)
                        combined_thk_1to5_layer3 = pd.concat([combined_thk_1to5_df_layer3, Temp_layer_create_1to5_df_layer3[2]])
                        # combined_thk_1to5_layer3.style.set_properties(**{'text-align': 'center'})
                        combined_thk_1to5_layer3.rename(columns={combined_thk_1to5_layer3.columns[3]:'Avg GOF'}, inplace=True)
                        combined_thk_1to5_layer3.loc[0,"Recipe"] = Thk_df.loc[0,"Jobfile Name"]
                        combined_thk_1to5_layer3.loc[:,"Lot ID"] = Thk_df.loc[:,"Lot ID"]
                        for i in range(len(No_of_slots)):
                            combined_thk_1to5_layer3.loc[i,"Slot No."] = No_of_slots[i]
                        Avg_thk_list_1to5_layer3 = []
                        for i in range(len(No_of_slots)):
                            Avg_thk_1to5_layer3 = combined_thk_1to5_layer3.iloc[i,4:].mean(skipna=True)
                            Avg_thk_list_1to5_layer3.append(Avg_thk_1to5_layer3)
                        combined_thk_1to5_layer3 = combined_thk_1to5_layer3.assign(Avg_Thk_5pts=Avg_thk_list_1to5_layer3)
                        ##print("This is combined_thk_1to5_layer3: ")
                        ##print(combined_thk_1to5_layer3)      
                        combined_thk_1to5_layer3["Avg GOF"] = combined_thk_1to5_layer3["Avg GOF"].round(decimals=3)
                        Thk_1to5_1decimal_layer3 = combined_thk_1to5_layer3.columns.get_loc("Avg GOF")
                        combined_thk_1to5_layer3.iloc[:,Thk_1to5_1decimal_layer3+1:] = combined_thk_1to5_layer3.iloc[:,Thk_1to5_1decimal_layer3+1:].round(decimals=1)

                        #All the dataframe placement done for both: 1) All pts and 2) 5 Pts. Time to put in excel file using xlwings (Because it can create temp workbook)

                        row_size_1to5 = int(combined_thk_1to5_layer3.shape[0])
                        col_size_1to5 = int(combined_thk_1to5_layer3.shape[1])

                        row_size_all_pts = int(combined_thk_all_pts_layer3.shape[0])
                        col_size_all_pts = int(combined_thk_all_pts_layer3.shape[1])

                        row_spacing_1to5_all_pts = int(combined_thk_1to5_layer3.shape[0]) + 2
                        col_spacing_1to5_all_pts = int(combined_thk_1to5_layer3.shape[1])

                        ws_wings3["B1"].value = "5 Points Thickness: "
                        ws_wings3["B1"].font.bold = True
                        ws_wings3.range(2,2).options(index=False, expand='table').value = combined_thk_1to5_layer3
                        ws_wings3.range((3,2),(2+row_size_1to5,2)).merge()
                        ws_wings3.range((2,2),(2+row_size_1to5,2+col_size_1to5-1)).api.Borders.Weight = 2
                        header_1to5_layer3 = ws_wings3.range(2,2).expand('right')
                        header_1to5_layer3.color = (209,252,237)
                        header_1to5_layer3.font.bold = True
                        ws_wings3.range(2+int(row_spacing_1to5_all_pts),2).value = "All Points Thickness: "
                        ws_wings3.range(2+int(row_spacing_1to5_all_pts),2).font.bold = True
                        ws_wings3.range(2+int(row_spacing_1to5_all_pts)+1,2).options(index=False).value = combined_thk_all_pts_layer3
                        ws_wings3.range((2+int(row_spacing_1to5_all_pts)+1+1,2),(2+int(row_spacing_1to5_all_pts)+1+row_size_all_pts,2)).merge()
                        ws_wings3.range(2+int(row_spacing_1to5_all_pts)+1+int(row_size_all_pts)+2,2).value = "Wafer Points on Map"
                        ws_wings3.range(2+int(row_spacing_1to5_all_pts)+1+int(row_size_all_pts)+2,2).font.bold = True
                        #Pivot_row1 = 2+int(row_spacing_1to5_all_pts)+1+int(row_size_all_pts)+2
                        #Pivot_col1 = 2                

                        header_all_pts_layer3 = ws_wings3.range(2+int(row_spacing_1to5_all_pts)+1,2).expand('right')
                        header_all_pts_layer3.color = (209,252,237)
                        header_all_pts_layer3.font.bold = True
                        ws_wings3.range((2+int(row_spacing_1to5_all_pts)+1,2),(2+int(row_spacing_1to5_all_pts)+1+row_size_all_pts,2+col_size_all_pts-1)).api.Borders.Weight = 2

                        #Pivot Table for Layer 3 Thk
                        #Create list of names to use for naming PIVOT TABLE based on no. of slots for LAYER 1
                        Pivot_slot_layer3 = []
                        for i in range(0, Max_no_of_slots):
                            string_layer3 = "PivotTable_Slot"+str(No_of_slots[i])
                            Pivot_slot_layer3.append(string_layer3)
                        #print(f"Pivot_slot_layer3: {Pivot_slot_layer3}")


                        #Create list of names for pivot table DATAFRAME based on no. of slots for LAYER1, to iterate during pivot table creation
                        Pivot_df_count_layer3 = 1
                        Pivot_df_layer3 = []
                        for i in range(0, Max_no_of_slots):
                            string_layer3 = "Layer_df_"+str(No_of_slots[i])
                            Pivot_df_layer3.append(string_layer3)
                            Pivot_df_count_layer3+=1

                        Max_site_count_layer3 = 0 
                        #Create pivot df with Thk_df "X Pos", "Y Pos", "Layer 2 Thickness"
                        for i in range(0, len(Pivot_df_layer3)):
                            Pivot_df_layer3[i] = Thk_df.loc[Max_site_count_layer3*Max_site:Max_site-1+Max_site_count_layer3*Max_site,["X Pos", "Y Pos", "Layer 3 Thickness"]]
                            #print(Pivot_df_layer3[i])
                            Max_site_count_layer3+=1

                        pivot_space_count_layer3 = 1
                        pivot_space_count2_layer3 = 0
                        for i in range(0, Max_no_of_slots):
                            Pivot_slot_layer3[i] = pd.pivot_table(Pivot_df_layer3[i], values="Layer 3 Thickness", index="Y Pos", columns="X Pos", aggfunc="mean")
                            row_size_pivot = Pivot_slot_layer3[i].shape[0]
                            col_size_pivot = Pivot_slot_layer3[i].shape[1]
                            #print(f"row_size_pivot = {row_size_pivot}")
                            ws_wings3.range(Pivot_row1+row_size_pivot*i+pivot_space_count_layer3+pivot_space_count2_layer3,Pivot_col1).value = Pivot_slot_layer3[i]
                            temp_pivot_header_col_layer3 = ws_wings3.range(Pivot_row1+row_size_pivot*i+pivot_space_count_layer3+pivot_space_count2_layer3,Pivot_col1).expand('right')
                            temp_pivot_header_col_layer3.color = (209,252,237)
                            temp_pivot_header_col_layer3.api.Borders.Weight=2
                            temp_pivot_header_row_layer3 = ws_wings3.range(Pivot_row1+row_size_pivot*i+pivot_space_count_layer3+pivot_space_count2_layer3,Pivot_col1).expand('down')
                            temp_pivot_header_row_layer3.color = (209,252,237)
                            temp_pivot_header_row_layer3.api.Borders.Weight=2

                            #Create color scale for pivot table values
                            Pivot_layer3_thk_list = Pivot_df_layer3[i]["Layer 3 Thickness"].tolist()
                            Pivot_layer3_thk_list_dcopy = copy.deepcopy(Pivot_layer3_thk_list)
                            Pivot_layer3_thk_list.sort()
                            ##print("list..")
                            ##print(Pivot_layer3_thk_list)
                            ##print(len(Pivot_layer3_thk_list))

                            Pivot_5pts_list_layer3 = []
                            Pivot_bold_list_layer3 = [True]*5
                            Pivot_5pts_dict_layer3 = {}

                            for l in range(0, len(Pivot_layer3_thk_list_dcopy), Max_site):
                                Pivot_5pts_list_layer3.append(Pivot_layer3_thk_list_dcopy[l:l+5])
                            Pivot_5pts_dict_layer3 = dict(zip(Pivot_5pts_list_layer3[0],Pivot_bold_list_layer3))

                            Red_div_layer3 = (255-0)/len(Pivot_layer3_thk_list)
                            Red_div_rounded_layer3 = np.floor(Red_div_layer3)
                            ##print(Red_div_rounded_layer3)

                            Blue_div_layer3 = (255-179)/len(Pivot_layer3_thk_list)
                            Blue_div_rounded_layer3 = np.floor(Blue_div_layer3)
                            ##print(Blue_div_rounded_layer3)

                            Pivot_color_scale_layer3 = []
                            for j in range(0,len(Pivot_layer3_thk_list)):
                                temp_pivot_tuple = (int(255-j*Red_div_rounded_layer3),int(255),int(255-j*Blue_div_rounded_layer3))
                                Pivot_color_scale_layer3.append(temp_pivot_tuple)
                            ##print(Pivot_color_scale_layer3)
                            
                            Pivot_color_scale_dict_layer3 = {}
                            Pivot_color_scale_dict_layer3 = dict(zip(Pivot_layer3_thk_list,Pivot_color_scale_layer3))
                            ##print(Pivot_color_scale_dict_layer3)

                            #Add red font to low GOF values in wafer map/pivot table
                            Low_GOF_pivot_red_font_dict_layer3 = {}
                            Low_GOF_pivot_red_font_dict_layer3 = dict(zip(Low_GOF_Thk_list[2],Low_GOF_red_font_list))                            

                            for k in ws_wings3.range((Pivot_row1+row_size_pivot*i+pivot_space_count_layer3+pivot_space_count2_layer3+1,Pivot_col1+1),(Pivot_row1+row_size_pivot*i+pivot_space_count_layer3+pivot_space_count2_layer3+row_size_pivot,Pivot_col1+col_size_pivot)):
                                if k.value is not None:
                                    k.color = Pivot_color_scale_dict_layer3.get(k.value)
                                    k.font.bold = Pivot_5pts_dict_layer3.get(k.value)
                                    try:
                                        k.font.color = Low_GOF_pivot_red_font_dict_layer3.get(k.value)
                                    except:
                                        pass
                                else:
                                    continue

                            ws_wings3.range(Pivot_row1+row_size_pivot*i+pivot_space_count_layer3+pivot_space_count2_layer3,Pivot_col1).value = "Slot "+ str(No_of_slots[i])
                            ws_wings3.range(Pivot_row1+row_size_pivot*i+pivot_space_count_layer3+pivot_space_count2_layer3,Pivot_col1).font.bold = True
                            pivot_space_count2_layer3+=1
                            pivot_space_count_layer3+=1
                        #End of pivot table + color scale function

                        Center_alignment_layer3 = ws_wings3.range((1,1),(Pivot_row1+row_size_pivot*Max_no_of_slots+20,1+col_size_all_pts+10))
                        Center_alignment_layer3.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
                        Center_alignment_layer3.api.VerticalAlignment = xw.constants.HAlign.xlHAlignCenter
                        ws_wings3.autofit(axis="columns")
                        App.thread_inserttext(app, text="...Layer 3 Done!")
                    
                    else:
                        
                        App.thread_inserttext(app, text="...No Layer 3 Detected")
                        
                else:
                    App.thread_inserttext(app, text="...No Layer 2 Detected")
                    App.thread_inserttext(app, text="...No Layer 3 Detected")
            App.thread_inserttext(app, text="All done for this file!\n")
        except:
            App.thread_inserttext(app, text="\n.....\nThere is an error =(\nPossible issues:\nTry going to Task Manager (Ctrl+Alt+Del) and cancel all Excel tasks and try again\nSame slot number repeated in raw data\nMissing column data\n...\n")       
            pass

    #End of Layer 3 Excel
    ##################################################################
        
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        
        self.title("Extract Data")
        self.geometry("578x473")
        self.resizable(0,0)
        self.emptylabel = Label(self, height=5, anchor="w")
        self.emptylabel.grid(row=2, column=0, sticky=E+W) 
        self.label = Label(self, text="Status of Data Extraction: ", background='#ECECEC', font=("Comic Sans", 11, "bold"))
        self.label.grid(row=3, column=0, columnspan=3, sticky=tk.EW, padx=5, pady=5) 
        self.ScrollScroll = scrolledtext.ScrolledText(self, wrap=tk.WORD, width=69, height=10, font=("Comic Sans",11))
        self.ScrollScroll.grid(row=4, column=0)
        self.ScrollScroll.insert(tk.INSERT,"Hi, welcome!\nMore will be added soon, feel free to leave me a feedback to improve this =)\n")
        self.create_widgets()
    
    def insert_scrolledtext(self, text):
        self.ScrollScroll.insert(tk.END,"\n"+str(text))
        self.ScrollScroll.see("end")

    def thread_inserttext(self, text):
        return Thread(target=self.insert_scrolledtext(text))

    def create_widgets(self):
        OT_frame = OT_Frame(self)
        OT_frame.grid(column=0, row=0) 
        Thk_frame = Thk_Frame(self)
        Thk_frame.grid(column=0, row=1)

if __name__ == "__main__":
    app = App()
    app.mainloop()
    