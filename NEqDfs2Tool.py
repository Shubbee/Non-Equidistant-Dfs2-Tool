"""
Created on Tue Jun  1 11:26:17 2021

Create Non-equidistant temporal axis Dfs2:

@author: Shubhneet Singh
ssin@dhigroup.com

"""

# marks dependencies
import os
import clr
import sys
import time
import datetime
from datetime import timedelta
import openpyxl
import numpy as np #
import pandas as pd #
import datetime as dt
import System
from System import Array

from winreg import ConnectRegistry, OpenKey, HKEY_LOCAL_MACHINE, QueryValueEx

def get_mike_bin_directory_from_registry():
    x86 = False
    dhiRegistry = "SOFTWARE\Wow6432Node\DHI\\"
    aReg = ConnectRegistry(None, HKEY_LOCAL_MACHINE)
    try:
        _ = OpenKey(aReg, dhiRegistry)
    except FileNotFoundError:
        x86 = True
        dhiRegistry = "SOFTWARE\Wow6432Node\DHI\\"
        aReg = ConnectRegistry(None, HKEY_LOCAL_MACHINE)
        try:
            _ = OpenKey(aReg, dhiRegistry)
        except FileNotFoundError:
            raise FileNotFoundError
    year = 2030
    while year > 2010:
        try:
            mikeHomeDirKey = OpenKey(aReg, dhiRegistry + str(year))
        except FileNotFoundError:
            year -= 1
            continue
        if year > 2020:
            mikeHomeDirKey = OpenKey(aReg, dhiRegistry + "MIKE Zero\\" + str(year))

        mikeBin = QueryValueEx(mikeHomeDirKey, "HomeDir")[0]
        mikeBin += "bin\\"

        if not x86:
            mikeBin += "x64\\"

        if not os.path.exists(mikeBin):
            print(f"Cannot find MIKE ZERO in {mikeBin}")
            raise NotADirectoryError
        return mikeBin

    print("Cannot find MIKE ZERO")
    return ""

sys.path.append(get_mike_bin_directory_from_registry())
clr.AddReference("DHI.Generic.MikeZero.DFS")
clr.AddReference("DHI.Generic.MikeZero.EUM")
clr.AddReference("DHI.Projections")

from DHI.Generic.MikeZero import eumUnit, eumItem, eumQuantity
from DHI.Generic.MikeZero.DFS import *
from DHI.Generic.MikeZero.DFS.dfs123 import *
from DHI.Projections import MapProjection

# import xlrd
from  mikeio import *
from  mikeio import Dfs2
from mikeio.eum import ItemInfo

from tkinter import Frame, Label, Button, Entry, Tk, W, END
from tkinter import messagebox as tkMessageBox
# from tkinter.filedialog import askdirectory
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfilename

def create_non_equidistant_calendar(dfs2file, data,
                                datetimes,
                                length_x = 1, length_y = 1,
                                x0 = 0, y0 = 0,
                                coordinate = None, variable_type=None, unit=None,
                                names=None, title=None):
    """
    Creates a dfs2 file
    dfs2file:
        Location to write the dfs2 file
    data:
        list of matrices, one for each item. Matrix dimension: y, x, time
    datetimes:
        list of datetimes
    variable_type:
        Array integers corresponding to a variable types (ie. Water Level). Use dfsutil type_list
        to figure out the integer corresponding to the variable.
    unit:
        Array integers corresponding to the unit corresponding to the variable types The unit (meters, seconds),
        use dfsutil unit_list to figure out the corresponding unit for the variable.
    coordinate:
        ['UTM-33', 12.4387, 55.2257, 327]  for UTM, Long, Lat, North to Y orientation. Note: long, lat in decimal degrees
    x0:
        Lower right position
    x0:
        Lower right position
    length_x:
        length of each grid in the x direction (meters)
    length_y:
        length of each grid in the y direction (meters)
    names:
        array of names (ie. array of strings). (can be blank)
    title:
        title of the dfs2 file (can be blank)
    """

    if title is None:
        title = ""

    number_y = np.shape(data[0])[0]
    number_x = np.shape(data[0])[1]
    n_time_steps = np.shape(data[0])[2]
    n_items = len(data)

    if coordinate is None:
        coordinate = ['LONG/LAT', 0, 0, 0]

    if names is None:
        names = [f"Item {i+1}" for i in range(n_items)]

    if variable_type is None:
        variable_type = [999] * n_items

    if unit is None:
        unit = [0] * n_items

    if not all( np.shape(d)[0] == number_y for d in data):
        raise Warning("ERROR data matrices in the Y dimension do not all match in the data list. "
                 "Data is list of matices [y,x,time]")
    if not all(np.shape(d)[1] == number_x for d in data):
        raise Warning("ERROR data matrices in the X dimension do not all match in the data list. "
                 "Data is list of matices [y,x,time]")
    if not all(np.shape(d)[2] == n_time_steps for d in data):
        raise Warning("ERROR data matrices in the time dimension do not all match in the data list. "
                 "Data is list of matices [y,x,time]")

    if not len(datetimes) == n_time_steps:
        raise Warning("Number of datetimes do not match number of time steps in time dimension")

    if len(names) != n_items:
        raise Warning("names must be an array of strings with the same number as matrices in data list")

    if len(variable_type) != n_items or not all(isinstance(item, int) and 0 <= item < 1e15 for item in variable_type):
        raise Warning("type if specified must be an array of integers (enuType) with the same number of "
                      "elements as data columns")

    if len(unit) != n_items or not all(isinstance(item, int) and 0 <= item < 1e15 for item in unit):
        raise Warning(
            "unit if specified must be an array of integers (enuType) with the same number of "
            "elements as data columns")


    start_time = datetimes[0]
    system_start_time = System.DateTime(start_time.year, start_time.month, start_time.day,
                                        start_time.hour, start_time.minute, start_time.second)

    # Create an empty dfs2 file object
    factory = DfsFactory();
    builder = Dfs2Builder.Create(title, 'pydhi', 0)

    # Set up the header
    builder.SetDataType(0)
    builder.SetGeographicalProjection(factory.CreateProjectionGeoOrigin(coordinate[0], coordinate[1], coordinate[2], coordinate[3]))
    builder.SetTemporalAxis(factory.CreateTemporalNonEqCalendarAxis(eumUnit.eumUsec, system_start_time))
    builder.SetSpatialAxis(factory.CreateAxisEqD2(eumUnit.eumUmeter, number_x, x0, length_x, number_y, y0, length_y))


    for i in range(n_items):
        builder.AddDynamicItem(names[i], eumQuantity.Create(variable_type[i], unit[i]), DfsSimpleType.Float, DataValueType.Instantaneous)

    try:
        builder.CreateFile(dfs2file)
    except IOError:
        print('cannot create dfs2 file: ', dfs2file)

    dfs = builder.GetFile();
    deletevalue = dfs.FileInfo.DeleteValueFloat #-1.0000000031710769e-30
    
    

    for i in range(n_time_steps):
        for item in range(n_items):
            d = data[item][:, :, i]
            d[np.isnan(d)] = deletevalue
            d = d.reshape(number_y, number_x)
            d = np.flipud(d)
            darray = Array[System.Single](np.array(d.reshape(d.size, 1)[:, 0]))
            t = datetimes[i]
            #sdt = (System.DateTime(t.year, t.month, t.day,
            #                            t.hour, t.minute, t.second) - system_start_time).TotalSeconds
            relt = (t-start_time).total_seconds()
            dfs.WriteItemTimeStepNext(relt, darray)

    dfs.Close()    
    
  

#------------------------------------------------------------------------------
# UI for this tool:
     
class interface(Frame):
    def __init__(self, master = None):
        """ Initialize Frame. """
        Frame.__init__(self,master)
        self.grid()
        self.createWidgets()
            
    def message(self):
        tkMessageBox.showinfo("Task Complete", "Dfs2 Created!")
    
    def run(self):
        
        # input1 - Data in excel:
        filename1 = self.file_name1.get()
        # filename1 = r"C:\Users\ssin\Downloads\LAI_1kmGRID_6_18_2021.dfs0"
        # input2 - grid cells:
        filename2 = self.file_name2.get()
        # filename2 = r"C:\Users\ssin\Downloads\VEGLAIcodes_USPR_1km_GRID_adjusted2.dfs2"
        # Output:
        outputFile = self.file_name6.get()
        
        # Tool
        begin_time = datetime.datetime.now()
        
              
        print('Reading input data...' )
        veg_code_file= Dfs2(filename2)
        veg_code_dfs2 = Dfs2(filename2).read()
        veg_timeseries_dfs0 = Dfs0(filename1).to_dataframe()
        print('Input data read. Output Dfs2 processing...' )
        print('Time taken: ' + str(round((datetime.datetime.now() - begin_time).total_seconds()/60,1)) + ' minutes')
        
        veg_code_array = veg_code_dfs2[0]

        veg_code_array_y = veg_code_array.shape[1]
        veg_code_array_x = veg_code_array.shape[2]
        
        
        Dfs2_Timesteps = [veg_timeseries_dfs0.index[i].round('1s') for i in range(len(veg_timeseries_dfs0.index))]       
        num_timesteps = len(veg_timeseries_dfs0)
        
        Dfs2_Data = np.empty((veg_code_array_y ,veg_code_array_x,num_timesteps))
        Dfs2_Data[:] = np.NaN
        
        for veg_code in veg_timeseries_dfs0.columns:
            
            for val_x in range(veg_code_array_x):
                for val_y in range(veg_code_array_y):
                    if float(veg_code) == veg_code_array[0][val_y][val_x]:
                        
                        Dfs2_Data[val_y, val_x, :] = veg_timeseries_dfs0[veg_code].values
          
        Dfs2_Data = [Dfs2_Data]
      
        if os.path.exists(outputFile):
            os.remove(outputFile)
        
        data_type =  [Dfs0(filename1).items[0].type]
        data_unit =  [Dfs0(filename1).items[0].unit]
        
        prj_str = DfsFileFactory.Dfs2FileOpen(filename2).FileInfo.Projection.WKTString
        orientation = DfsFileFactory.Dfs2FileOpen(filename2).FileInfo.Projection.Orientation
        long = DfsFileFactory.Dfs2FileOpen(filename2).FileInfo.Projection.Longitude
        lat = DfsFileFactory.Dfs2FileOpen(filename2).FileInfo.Projection.Latitude
        coordinate_input = [prj_str, long, lat, orientation]
        
        print('Writing output Dfs2...' )
        create_non_equidistant_calendar(dfs2file = outputFile,
                                        data = Dfs2_Data,
                                        datetimes = Dfs2_Timesteps,
                                        length_x = veg_code_file.dx,
                                        length_y = veg_code_file.dy,
                                        x0 = 0, y0 = 0,
                                        coordinate = coordinate_input,
                                        variable_type= data_type,
                                        unit = data_unit,
                                        names= ['Leaf Area Index'],
                                        title='LAI')       
        print('Dfs2 created with ' +  str(num_timesteps) + ' Non-Equidistant time-steps' )
        print('Total time taken: ' + str(round((datetime.datetime.now() - begin_time).total_seconds()/60,1)) + ' minutes')
        self.message()
        

    def createWidgets(self):
        
        # set all labels of inputs:

        Label(self, text = "LAI Data (.dfs0) :")\
            .grid(row=0, column=0, sticky=W)
        Label(self, text = "LAI Grid Codes (.dfs2) :")\
            .grid(row=1, column=0, sticky=W)            
        Label(self, text = "Output File (.dfs2) :")\
            .grid(row=2, column=0, sticky=W)
         
            
        # set buttons
        Button(self, text = "Browse", command=self.load_file1, width=10)\
            .grid(row=0, column=6, sticky=W)
        Button(self, text = "Browse", command=self.load_file2, width=10)\
            .grid(row=1, column=6, sticky=W)
            
        Button(self, text = "Save As", command=self.load_file6, width=10)\
            .grid(row=2, column=6, sticky=W)            
        Button(self, text = "Run", command=self.run, width=20)\
            .grid(row=3, column=3, sticky=W)
       
        # set entry field
        self.file_name1 = Entry(self, width=65)
        self.file_name1.grid(row=0, column=1, columnspan=4, sticky=W)
        
        self.file_name2 = Entry(self, width=65)
        self.file_name2.grid(row=1, column=1, columnspan=4, sticky=W)
        
        self.file_name6 = Entry(self, width=65)
        self.file_name6.grid(row=2, column=1, columnspan=4, sticky=W)


    def load_file1(self):
        self.filename = askopenfilename(initialdir=os.path.curdir, defaultextension=".dfs0", filetypes=(("Dfs0 File", "*.dfs0"),("All Files", "*.*") ))
        if self.filename: 
            try: 
                #self.settings.set(self.filename)
                self.file_name1.delete(0, END)
                self.file_name1.insert(0, self.filename)
                self.file_name1.xview_moveto(1.0)
            except IOError:
                tkMessageBox.showerror("Error","Failed to read file \n'%s'"%self.filename) 

    def load_file2(self):
        self.filename = askopenfilename(initialdir=os.path.curdir, defaultextension=".dfs2", filetypes=(("Dfs2 File", "*.dfs2"),("All Files", "*.*") ))
        if self.filename: 
            try: 
                #self.settings.set(self.filename)
                self.file_name2.delete(0, END)
                self.file_name2.insert(0, self.filename)
                self.file_name2.xview_moveto(1.0)
            except IOError:
                tkMessageBox.showerror("Error","Failed to read file \n'%s'"%self.filename)                 
   
    def load_file6(self):
        self.filename = asksaveasfilename(initialdir=os.path.curdir,defaultextension=".dfs2", filetypes=(("Dfs2 File", "*.dfs2"),("All Files", "*.*") ))
        if self.filename: 
            try: 
                #self.settings.set(self.filename)
                self.file_name6.delete(0, END)
                self.file_name6.insert(0, self.filename)
                self.file_name6.xview_moveto(1.0)
            except IOError:
                tkMessageBox.showerror("Error","Failed to read file \n'%s'"%self.filename) 
                
##### main program

root = Tk()
UI = interface(master=root)
UI.master.title("Non-Equidistant Dfs2 Tool")
UI.master.geometry('630x160')
for child in UI.winfo_children():
    child.grid_configure(padx=4, pady =6)

file_name2 = Entry(root)

    
UI.mainloop()
