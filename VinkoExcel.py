#!/usr/bin/env python
# coding: utf-8

# In[1]:


from tkinter import *
import tkinter as tk
import tkinter.messagebox
from datetime import datetime
from tkinter.ttk import *
import os
import re
import pandas as pd
import numpy 
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

'''Crating a list of all excel files in the folder'''
Direc = os.getcwd()
files = os.listdir(Direc)
files = [f for f in files if os.path.isfile(Direc+'/'+f)] #Filtering only the files.
r = re.compile(".*xlsx")
file_options = list(filter(r.match, files))
options_for_user = ['remove column', 'add column', 'rename column', 'swap column', 'export file', 'data visualization']
options_for_column_to_add = ['fill 0', 'sum', 'average']
options_for_data_visualization = ['plot']
track_list = []

class VinkoExcel(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)
        self.geometry("800x400")
        self.title('VinkoExcel')
       # Label(self, text = 'Welcome to VinkoExcel!').grid()
        '''Creating frames'''
        #Creating Buttonframe
        self.buttonframe = Frame(self, width = 100, height = 50)
        self.buttonframe.grid()
        
        # variables for entry
        self.rename_column_var = StringVar()
        self.file_name_var = StringVar()
        self.add_column_var = StringVar()
        self.print_text_var = StringVar()
        
        #flags
        self.flag_chart_up = FALSE
        
        #Buttons
        self.launch = Button(self.buttonframe, text = 'Launch', command = self.Launch)
        self.select_file = Button(self.buttonframe, text = 'Select file', command = self.SelectFile)
        self.select_sheet = Button(self.buttonframe, text = 'Select sheet', command = self.SelectSheet)
        self.clear_table = Button(self.buttonframe, text = 'Clear table', command = self.clear_treeview)
        #self.swap_columns = Button(self.buttonframe, text = 'Swap Columns', command = self.swapColumns)
        self.execute_swap_columns = Button(self.buttonframe, text = 'Execute Swap Columns', command = self.execute_swaping_columns)
        self.option_manipulate_column = Button(self.buttonframe, text = 'Chose option:', command = self.selectingOptions)
        self.b_remove_column = Button(self.buttonframe, text = 'Remove Column', command = self.removeColumn)
        self.b_rename_column = Button(self.buttonframe, text = 'Rename Column', command = self.executeColumnRename)
        self.b_export_column = Button(self.buttonframe, text = 'Export', command = self.ExportFile)
        self.b_add_column = Button(self.buttonframe, text = 'Add Column', command = self.AddColumn)
        self.b_summ_column = Button(self.buttonframe, text = 'Sum Column', command = self.SumColumn)
        self.b_data_visualization = Button(self.buttonframe, text = 'Select', command = self.SelectingChartType)
        self.b_create_plot = Button(self.buttonframe, text = 'Create Chart', command = self.ShowPlotChart)
        self.b_save_chart_as_png = Button(self.buttonframe, text = 'Chart to png', command = self.PrintChart)
        self.b_back = Button(self.buttonframe, text = 'Back', command = self.Back)
        
        
        #Labels
        self.lb_chose_column_to_delete = Label(self.buttonframe,text = 'Chose column to to be deleted!')
        self.lb_first_column_swap = Label(self.buttonframe,text = 'Chose first column to swap!')
        self.lb_second_column_swap = Label(self.buttonframe,text = 'Chose second column to swap!')
        self.lb_column_to_rename = Label(self.buttonframe,text = 'Chose column to rename!')
        self.lb_rename_to = Label(self.buttonframe,text = 'Rename to:')
        self.lb_file_name = Label(self.buttonframe,text = 'Enter the name of your file:')
        self.lb_add_column = Label(self.buttonframe,text = 'Enter name of the column:')
        self.lb_add_column_option = Label(self.buttonframe,text = 'Options for adding:')
        self.lb_sum_column_option = Label(self.buttonframe,text = 'Which columns to sum:')
        self.lb_data_visualization = Label(self.buttonframe,text = 'Select chart type:')
        self.lb_select_xaxis = Label(self.buttonframe,text = 'x-axis data selection')
        self.lb_select_yaxis = Label(self.buttonframe,text = 'y-axis data selection')
        
        
        #ListBoxes
        self.option_box = Listbox(self.buttonframe)
        self.sheets_box = Listbox(self.buttonframe)
        self.swap_first_column = Listbox(self.buttonframe)
        self.swap_second_column = Listbox(self.buttonframe, exportselection=False)
        self.remove_column = Listbox(self.buttonframe)
        self.lbo_rename_column = Listbox(self.buttonframe)
        self.lbo_add_column_option = Listbox(self.buttonframe)
        self.lbo_summ_column = Listbox(self.buttonframe, selectmode = "multiple")
        self.lbo_visualization_option = Listbox(self.buttonframe)
        self.lbo_xaxis_options = Listbox(self.buttonframe)
        self.lbo_yaxis_options = Listbox(self.buttonframe, exportselection=False)
        
        #Entries
        self.ent_rename_column = Entry(self.buttonframe,textvariable = self.rename_column_var)
        self.ent_file_name = Entry(self.buttonframe,textvariable = self.file_name_var)
        self.ent_add_column = Entry(self.buttonframe,textvariable = self.add_column_var)
        self.ent_print = Entry(self.buttonframe,textvariable = self.print_text_var)
        
        #ComboBoxes
        self.list_of_options = Combobox(state = 'readonly', values = options_for_user )
        #self.cmbo_add_column_option = Combobox(state = 'readonly', values = options_for_column_to_add)

                
        #Button positioning
        self.launch.grid(row = 1)
        self.b_back.grid(row = 1, column = 6)

        
    def clear_treeview(self):
        '''clearing the dataframe data'''
        self.tree.delete(*self.tree.get_children())
    
    def Launch(self):
        '''starting the program and creating selectfile button and the option box'''
        #self.option_box.delete(0,'end') # removing everything  in file options
        self.select_file.grid(row = 1)
        Label(self.buttonframe,text = 'Chose file.').grid(row = 2, column = 0)
        self.option_box.grid(row = 3, column = 0)
        self.launch.grid_forget()
        for file in range(len(file_options)):
            self.option_box.insert('end', file_options[file])
        self.openTableWindow()
    
    def SelectFile(self):
        '''selecting the working file and then creating the selectsheet button and the sheet option box'''
        self.sheets_box.delete(0,'end')
        Label(self.buttonframe,text = 'Chose sheet.').grid(row = 2, column = 1)
        #self.show.config(text=self.option_box.get(ANCHOR))
        self.chosen_file = self.option_box.get(ANCHOR)
        xl = pd.ExcelFile(self.chosen_file)  #picking sheets
        sheets_list = xl.sheet_names
        self.sheets_box.grid(row = 3, column = 1)
        self.select_sheet.grid(row = 1, column = 1)
        for sheet in range(len(sheets_list)):
            self.sheets_box.insert('end',sheets_list[sheet])
        
    def SelectSheet(self):
        '''selecting the sheet we want to work with'''
        self.chosen_sheet = self.sheets_box.get(ANCHOR)
        self.df = pd.read_excel(self.chosen_file, sheet_name = self.chosen_sheet)
        '''filling the treeview with data from the df we chose'''
        self.tree["column"] = list(self.df.columns)
        self.tree["show"] = "headings"
        self.df_rows = self.df.to_numpy().tolist()
        
        self.SaveVersionDf()
        self.loading_table()
        self.option_manipulate_column.grid(row = 1, column = 7)
        #Label(self.buttonframe, text = 'Options:').grid(row = 2, column = 6)
        self.list_of_options.grid(row = 0 , column = 7, sticky = NE)
        self.clear_table.grid(row = 1, column = 3)
        
    
    def swapColumns(self):
        '''Popup the two list boxes through which the user can select columns'''
        self.lb_first_column_swap.grid(row = 2, column = 4)
        self.lb_second_column_swap.grid(row = 2, column = 5)
        self.swap_first_column.grid(row =3, column = 4)
        self.swap_second_column.grid(row =3, column = 5)
        self.execute_swap_columns.grid(row = 1, column = 4)
        self.first_column_swap = self.swap_first_column.get(ANCHOR)
        self.second_column_swap = self.swap_second_column.get(ANCHOR)
            
    def execute_swaping_columns(self):
        '''Swaping the two columns, chosen by the user'''
        self.first_column_swap = self.swap_first_column.get(ANCHOR)
        self.second_column_swap = self.swap_second_column.get(ANCHOR)
        first_column_index = self.tree["column"].index(self.first_column_swap)
        second_column_index = self.tree["column"].index(self.second_column_swap)
        moved_column_one = self.df.pop(self.first_column_swap)
        moved_column_two = self.df.pop(self.second_column_swap)
        
        if first_column_index < second_column_index:
            self.df.insert(first_column_index, self.second_column_swap, moved_column_two)
            self.df.insert(second_column_index, self.first_column_swap, moved_column_one)
        else:
            self.df.insert(second_column_index, self.first_column_swap, moved_column_one)
            self.df.insert(first_column_index, self.second_column_swap, moved_column_two)
        
        self.UpdatingListboxes()
        
    
    def loading_table(self):
        '''Loads the table with data'''
        self.clear_treeview()
        # For Headings iterate over the columns
        for col in self.tree["column"]:
            self.tree.heading(col, text=col)
            self.tree.column(col, stretch=True, width = 50)
        # Put Data in Rows
        for row in self.df_rows:
            self.tree.insert("", "end", values=row)
            self.tree.pack(fill = X)
        
        #Creating a scrollbar
        self.vertical_frame.config(command = self.tree.yview)
        self.vertical_frame.pack(side = RIGHT, fill = Y)
        self.tree['yscrollcommand'] = self.vertical_frame
        
        self.horizontal_frame.config(command = self.tree.xview)
        #self.horizontal_frame.pack(side = BOTTOM, fill = X)
        self.tree['xscrollcommand'] = self.horizontal_frame
        
    def openTableWindow(self):
        '''Opening another window, where the data will be displayed'''
        self.tableWindow = Toplevel(self)
        self.tableWindow.geometry('500x300')
        self.tableWindow.title('Table Window')
        self.vertical_frame = Scrollbar(self.tableWindow, orient="vertical")
        self.horizontal_frame = Scrollbar(self.tableWindow, orient="horizontal")
        
        self.tree = Treeview(self.tableWindow, h = 10, xscrollcommand=self.horizontal_frame.set, yscrollcommand=self.vertical_frame.set)
        #clear_treeview()
        
        
    def PrepareForRemove(self):
        '''Setting the layout for the remove option'''
        self.lb_chose_column_to_delete.grid(row = 2, column = 4)
        self.remove_column.grid(row = 3, column = 4)
        self.b_remove_column.grid(row = 1, column = 4)
            
    
    def removeColumn(self):
        '''Removes the chosen column by the user'''
        column_to_remove = self.remove_column.get(ANCHOR)
        self.df = self.df.drop(column_to_remove, axis=1)
        
        self.SaveVersionDf()
        self.UpdatingListboxes()
        
    def selectingOptions(self):
        '''Operates which option you've chose and calls the function, that is bind to it'''
        self.hideButtonsOptions()
        self.UpdatingListboxes()
        if self.list_of_options.get() == 'remove column':
            self.PrepareForRemove()
                
        elif self.list_of_options.get() == 'swap column':
            self.swapColumns()
        
        elif self.list_of_options.get() == 'rename column':
            self.renameColumn()
            
        elif self.list_of_options.get() == 'export file':
            self.b_export_column.grid(row = 1, column = 4)
            self.lb_file_name.grid(row = 2, column = 4)
            self.ent_file_name.grid(row = 3, column = 4, sticky = N)
            
        elif self.list_of_options.get() == 'add column':
            self.PrepAddColumn()
            
        elif self.list_of_options.get() == 'data visualization':
            self.PrepDataVisualization()
            
    def hideButtonsOptions(self):
        '''Removing anything from previous chosen option'''
        # Remove buttons
        self.remove_column.grid_forget()
        self.b_remove_column.grid_forget()
        self.swap_first_column.grid_forget()
        self.swap_second_column.grid_forget()
        self.execute_swap_columns.grid_forget()
        self.b_rename_column.grid_forget()
        self.b_export_column.grid_forget()
        self.b_add_column.grid_forget()
        self.b_summ_column.grid_forget()
        self.b_data_visualization.grid_forget()
        self.b_create_plot.grid_forget()
        self.b_save_chart_as_png.grid_forget()
        
        #Remove Labels
        self.lb_column_to_rename.grid_forget()
        self.lb_rename_to.grid_forget()
        self.lb_first_column_swap.grid_remove()
        self.lb_second_column_swap.grid_forget()
        self.lb_chose_column_to_delete.grid_remove()
        self.lb_file_name.grid_forget()
        self.lb_add_column.grid_forget()
        self.lb_add_column_option.grid_forget()
        self.lb_sum_column_option.grid_forget()
        self.lb_data_visualization.grid_forget()
        self.lb_select_xaxis.grid_forget()
        self.lb_select_yaxis.grid_forget()
        
        #Remove Listbox
        self.lbo_rename_column.grid_forget()
        self.lbo_add_column_option.grid_forget()
        self.lbo_summ_column.grid_forget()
        self.lbo_visualization_option.grid_forget()
        self.lbo_xaxis_options.grid_forget()
        self.lbo_yaxis_options.grid_forget()
        
        #Remove Entries
        self.ent_rename_column.grid_forget()
        self.ent_file_name.grid_forget()
        self.ent_add_column.grid_forget()
        
        #Remove Combobox

    
    def renameColumn(self):
        '''Calling everything you need for the executeColumnRename'''
        self.lbo_rename_column.grid(row = 3, column = 4)
        self.b_rename_column.grid(row = 1, column = 4)
        self.lb_column_to_rename.grid(row = 2, column = 4)
        self.lb_rename_to.grid(row = 2, column = 5)
        self.ent_rename_column.grid(row = 3, column = 5, sticky = N)
        
        
            
        
    def executeColumnRename(self):
        '''By pressing the button Rename and writing to what you want to change it, the chosen column  will be changed'''
        rename_to_column = self.rename_column_var.get()
        chosen_column_to_rename = self.lbo_rename_column.get(ANCHOR)
        if rename_to_column not in self.tree["column"]:
            self.df = self.df.rename(columns={chosen_column_to_rename: rename_to_column})
        else:
            tkinter.messagebox.showinfo("Error",  "Column with this name already exists!")
            
        self.SaveVersionDf()    
        self.UpdatingListboxes()
        
    def ExportFile(self):
        '''Exports the xlsx file with a name of your choice'''
        file_name = self.file_name_var.get() + '.xlsx'
        if file_name not in file_options:
            self.df.to_excel(f"{file_name}",index=False)
            tkinter.messagebox.showinfo("Done",  "Your file has been created!")
        else:
            tkinter.messagebox.showinfo("Error",  "File with this name already exists!")
            
    def UpdatingListboxes(self):
        '''Updating the information in the Listboxes'''
        self.tree["column"] = list(self.df.columns)
            
        #rename 
        self.lbo_rename_column.delete(0,'end')
        for col in self.tree["column"]:
            self.lbo_rename_column.insert('end', col)
        
        #remove
        self.remove_column.delete(0,'end')
        for col in self.tree["column"]:
            self.remove_column.insert('end', col)
        
        #swap
        self.swap_first_column.delete(0,'end')
        self.swap_second_column.delete(0,'end')
        for col in self.tree["column"]:
            self.swap_first_column.insert('end', col)
        for col in self.tree["column"]:
            self.swap_second_column.insert('end', col)
            
        #add
        self.lbo_summ_column.delete(0, 'end')
        for col in self.tree["column"]:
            self.lbo_summ_column.insert('end', col)
            
        self.df_rows = self.df.to_numpy().tolist()
        self.loading_table()
    
        
        
    def PrepAddColumn(self):
        '''Add the buttons and labels in the layout to add a column'''
        self.b_add_column.grid(row = 1, column = 4)
        self.lb_add_column.grid(row = 2, column = 4)
        self.ent_add_column.grid(row = 3, column = 4, sticky = N)
        self.lb_add_column_option.grid(row = 2, column = 5)
        self.lbo_add_column_option.grid(row = 3, column = 5)
        self.lbo_add_column_option.delete(0, 'end')
        for option in options_for_column_to_add:
            self.lbo_add_column_option.insert('end', option)
    
    def AddColumn(self):
        '''Create an empty column (fill) or create a column with values =0'''
        self.add_column_name = self.add_column_var.get()
        if self.lbo_add_column_option.get(ANCHOR) == 'fill 0':
            self.df[self.add_column_name] = np.nan

        elif self.lbo_add_column_option.get(ANCHOR) == 'sum':
            self.lbo_summ_column.grid(row = 3, column = 6)
            self.b_summ_column.grid(row = 1, column = 5)
            self.lb_sum_column_option.grid(row = 2, column = 6)
            self.lbo_summ_column.delete(0,'end')
            self.df[self.add_column_name] = 0
            for col in self.tree["column"]:
                self.lbo_summ_column.insert('end', col)
                
        #self.SaveVersionDf()
        self.UpdatingListboxes()
    
    def SumColumn(self):
        '''Sum of chosen columns, which contain numeric values'''
        to_sum_column = self.lbo_summ_column.curselection()
        for item in to_sum_column:
            if self.df[self.tree["column"][item]].dtypes != object:
                self.df[self.add_column_name] += self.df[self.tree["column"][item]]
                
        #self.SaveVersionDf()
        self.UpdatingListboxes() 
        
    def PrepDataVisualization(self):
        '''Add the buttons and labels in the layout to data visualization'''
        self.b_data_visualization.grid(row = 1, column = 4)
        self.lb_data_visualization.grid(row = 2, column = 4)
        self.lbo_visualization_option.grid(row = 3, column = 4)
        self.lbo_visualization_option.delete(0, 'end')
        for option in options_for_data_visualization:
            self.lbo_visualization_option.insert('end', option)
        
        self.UpdatingListboxes()
        #self.SelectingChartType()
            
    def SelectingChartType(self):
        '''Create ListBoxes so that x and y axis columns are selected and create the selected chart type'''
        if self.lbo_visualization_option.get(ANCHOR) == 'plot':
            self.CreatePlotChart()
            
        #self.hideButtonsOptions()
        
    def CreatePlotChart(self):
        '''Create the plot chart based on the users selection for the x and y axes'''
        self.b_create_plot.grid(row = 1, column = 4)
        self.b_save_chart_as_png.grid(row = 1, column = 5)
        self.lb_select_xaxis.grid(row = 2, column = 4)
        self.lb_select_yaxis.grid(row = 2, column = 5)
        self.lbo_xaxis_options.grid(row = 3, column = 4)
        self.lbo_yaxis_options.grid(row = 3, column = 5)
        
        self.lbo_xaxis_options.delete(0,'end')
        for col in self.tree["column"]:
            self.lbo_xaxis_options.insert('end', col)
        self.lbo_yaxis_options.delete(0,'end')
        for col in self.tree["column"]:
            self.lbo_yaxis_options.insert('end', col)
            
    def ShowPlotChart(self):
        '''Create and visualize the diagram'''
        chosen_xaxis = self.lbo_xaxis_options.get(ANCHOR)
        chosen_yaxis = self.lbo_yaxis_options.get(ANCHOR)
        xaxis = np.array(self.df[chosen_xaxis])
        yaxis = np.array(self.df[chosen_yaxis])
        #plt.plot(chosen_xaxis, chosen_yaxis, data = self.df)
        plot_chart = self.df[[chosen_xaxis, chosen_yaxis]].head(20).groupby(chosen_xaxis).sum()
        
        if self.flag_chart_up:
            self.line1.get_tk_widget().destroy()
        
        #Creating plot element in the toplevel
        self.figure_plot = plt.Figure(figsize=(10, 10), dpi=100)
        ax1 = self.figure_plot.add_subplot(111)
        self.line1 = FigureCanvasTkAgg(self.figure_plot, self.tableWindow)
        self.line1.get_tk_widget().pack(fill=tk.BOTH)
    
        plot_chart.plot(kind='bar', legend=True, ax=ax1)
        ax1.set_title('Pray to Jesus')
        ax1.set_xlabel(chosen_xaxis)
        
        self.flag_chart_up = TRUE
        
    def PrintChart(self):
        '''Create png from chart'''
        self.figure_plot.savefig('chart.png')
        tkinter.messagebox.showinfo("Done",  "Your chart is saved!")
    
    def Back(self):
        '''Takes the previous version of the df'''
        if len(track_list) >= 2:
            self.df = track_list[len(track_list) - 2]
            track_list.pop()
        else:
            tkinter.messagebox.showinfo("!",  "This is the first version")
        self.UpdatingListboxes()
            
    def SaveVersionDf(self):
        '''Prepares a list of all df versions'''
        track_list.append(self.df)
        
        
            
app = VinkoExcel()
app.mainloop() 


# In[ ]:




