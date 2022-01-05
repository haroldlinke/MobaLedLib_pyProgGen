# -*- coding: utf-8 -*-
#
#         Write header
#
# * Version: 1.21
# * Author: Harold Linke
# * Date: January 1st, 2020
# * Copyright: Harold Linke 2020
# *
# *
# * MobaLedCheckColors on Github: https://github.com/haroldlinke/MobaLedCheckColors
# *
# *
# * History of Change
# * V1.00 10.03.2020 - Harold Linke - first release
# *  
# * https://github.com/Hardi-St/MobaLedLib
# *
# * MobaLedCheckColors is free software: you can redistribute it and/or modify
# * it under the terms of the GNU General Public License as published by
# * the Free Software Foundation, either version 3 of the License, or
# * (at your option) any later version.
# *
# * MobaLedCheckColors is distributed in the hope that it will be useful,
# * but WITHOUT ANY WARRANTY; without even the implied warranty of
# * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# * GNU General Public License for more details.
# *
# * You should have received a copy of the GNU General Public License
# * along with this program.  if not, see <http://www.gnu.org/licenses/>.
# *
# *
# ***************************************************************************

from tkintertable import TableCanvas, TableModel
import tkinter as tk
from tkinter import ttk
from mlpyproggen.X01_Excel_Consts import *
from vb2py.vbconstants import *
from vb2py.vbfunctions import *
#import mlpyproggen.M20_PageEvents_a_Functions as M20


def TimeValue(Duration):
    
    return 5

def ActiveCell():
    global Selection
    table = ActiveSheet.table
    row = table.getSelectedRow()+1
    col = table.getSelectedColumn()+1
    Selection = CSelection(ActiveSheet.Cells(row,col)) #*HL
    return ActiveSheet.Cells(row,col) #*HL

def create_workbook(frame=None, path=None):
    global ThisWorkbook,ActiveSheet,ActiveWorkbook
    ActiveWorkbook = Workbook = ThisWorkbook = CWorkbook(frame=frame,path=path)
    ActiveSheet = ActiveWorkbook.Sheets("DCC")
    return ActiveWorkbook
    
def IsError(testval):
    return False
    

def Cells(row:int, column:int):
    cell = ActiveSheet.Cells(row,column)
    return cell

def Range(cell1,cell2):
    return CRange(cell1,cell2)

def Sheets(sheetname):
    return ActiveWorkbook.Sheets(sheetname)

def Rows(row:int):
    sheet = ActiveSheet
    return sheet.MaxRange.Rows[row]
   
def Columns(col:int):
    sheet = ActiveSheet
    return sheet.Maxrange.Columns[col]

def IsEmpty(obj):
    if len(obj)==0:
        return True
    else:
        return False

def val(value):
    valtype=type(value)
    if valtype is CCell:
        cellvalue = value.get_value()
        if cellvalue != "":
            return int(cellvalue)
        else:
            return 0
    elif valtype is str:
        if value != "":
            if IsNumeric(value):
                return int(value)
            else:
                return 0
        else:
            return 0
    return int(value)

def VarType(obj):
    valtype=type(obj)
    if valtype is str:
        return vbString
    else:
        return None

def ChDrive(srcdir):
    print("ChDrive:", srcdir)
    return

def Format(value,formatstring):
    return str(value) # no formating implemented yet


    

def MsgBox(ErrorMessage:str, msg_type:int, ErrorTitle:str):
    if msg_type == vbQuestion + vbYesNoCancel:
        res=tk.messagebox.askyesnocancel(title=ErrorTitle, message=ErrorMessage)
        if res == None:
            return vbCancel
        if res:
            return vbYes
        else:
            return vbNo
    elif msg_type == vbOKCancel:
        res=tk.messagebox.askokcancel(title=ErrorTitle, message=ErrorMessage)
        if res == None:
            return vbCancel
        if res:
            return vbOK
        else:
            return vbCancel
    
    else:
        res=tk.messagebox.askyesnocancel(title=ErrorTitle, message=ErrorMessage)
        if res == None:
            return vbCancel
        if res:
            return vbYes
        
    return vbNo

def InputBox(Message:str, Title:str, Default=None):
    res = tk.simpledialog.askstring(Title,Message,initialvalue=Default)
    return res


class CWorkbook:
    def __init__(self, frame=None,path=None):
        global Workbooks
        # Row and Columns are 0 based and not 1 based as in Excel
        sheetdict={"DCC":
                    {"Name":"DCC",
                     "Filename"  : "\\csv\\dcc.csv",
                     "Fieldnames": ("A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q"),
                     "Formating" : { "HideCells"      : ((0,"*"),),
                                    "ProtectedCells"  : ((0,"*"),(1,"*"), ("*",12),("*",13),("*",14),("*",15),("*",16)),
                                    "FontColor"       : { "1": {
                                                                "font"     : ("Arial",10),
                                                                "fg"       : "#FFFF00",
                                                                "bg"       : "#0000FF",
                                                                "Cells"    : ((1,"*"),)
                                                                },
                                                          "2": {
                                                                "font"     : ("Wingdings",10),
                                                                "fg"       : "#000000",
                                                                "bg"       : "#FFFFFF",
                                                                "Cells"    : (("*",1),)
                                                                },                                                          
                                                        "default": {
                                                               "font"     : ("Arial",10),
                                                               "fg"       : "#000000",
                                                               "bg"       : "#FFFFFF",
                                                               "Cells"    : (("*","*"),)
                                                               }
                                                        }
                        }
                     },
                   "Selectrix":
                                       {"Name":"Selectrix",
                                        "Filename"  : "\\csv\\Selectrix.csv",
                                        "Fieldnames": ("A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q"),
                                        "Formating" : { "HideCells"       : ((0,0),(0,1),(0,2),(0,3),(0,4),(0,5),(0,6),(0,7),(0,8),(0,9),(0,10),(0,11),(0,12),(0,13),(0,14)),
                                                       "ProtectedCells"  : ((0,"*"),(1,"*"), ("*",12),("*",13),("*",14),("*",15),("*",16)),
                                                       "FontColor"       : { "1": {
                                                                                   "font"     : ("Arial",10),
                                                                                   "fg"       : "#FFFF00",
                                                                                   "bg"       : "#0000FF",
                                                                                   "Cells"    : ((1,"*"),)
                                                                                   },
                                                                             "2": {
                                                                                   "font"     : ("Wingdings",10),
                                                                                   "fg"       : "#000000",
                                                                                   "bg"       : "#FFFFFF",
                                                                                   "Cells"    : (("*",1),)
                                                                                   },                                                          
                                                                           "default": {
                                                                                  "font"     : ("Arial",10),
                                                                                  "fg"       : "#000000",
                                                                                  "bg"       : "#FFFFFF",
                                                                                  "Cells"    : (("*","*"),)
                                                                                  }
                                                                           }
                                           }
                                        },
                   "CAN":
                                       {"Name":"CAN",
                                        "Filename"  : "\\csv\\CAN.csv",
                                        "Fieldnames": ("A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q"),
                                        "Formating" : { "HideCells"       : ((0,0),(0,1),(0,2),(0,3),(0,4),(0,5),(0,6),(0,7),(0,8),(0,9),(0,10),(0,11),(0,12),(0,13),(0,14)),
                                                       "ProtectedCells"  : ((0,"*"),(1,"*"), ("*",12),("*",13),("*",14),("*",15),("*",16)),
                                                       "FontColor"       : { "1": {
                                                                                   "font"     : ("Arial",10),
                                                                                   "fg"       : "#FFFF00",
                                                                                   "bg"       : "#0000FF",
                                                                                   "Cells"    : ((1,"*"),)
                                                                                   },
                                                                             "2": {
                                                                                   "font"     : ("Wingdings",10),
                                                                                   "fg"       : "#000000",
                                                                                   "bg"       : "#FFFFFF",
                                                                                   "Cells"    : (("*",1),)
                                                                                   },                                                          
                                                                           "default": {
                                                                                  "font"     : ("Arial",10),
                                                                                  "fg"       : "#000000",
                                                                                  "bg"       : "#FFFFFF",
                                                                                  "Cells"    : (("*","*"),)
                                                                                  }
                                                                           }
                                           }
                                        },                   
                   "Config":
                    {"Name":"Config",
                     "Filename":"\\csv\\Config.csv",
                     "Fieldnames": "A;B;C;D",
                     "Formating" : { "HideCells"       : (("*",3),),
                                    "ProtectedCells"  : ( ("*",1),("*",3)),
                                    "FontColor"       : { "1": {
                                                                "font"     : ("Arial",10),
                                                                "fg"       : "#0000FF",
                                                                "bg"       : "#FFFFFF",
                                                                "Cells"    : ((0,"*"),)
                                                                },
                                                        "default": {
                                                               "font"     : ("Arial",10),
                                                               "fg"       : "#000000",
                                                               "bg"       : "#FFFFFF",
                                                               "Cells"    : (("*","*"),)
                                                               }
                                                        }
                                    }
                    },
                   "Languages":
                    {"Name":"Languages",
                     "Filename":"\\csv\\Languages.csv",
                     "Fieldnames": "A;B;C;D;E;F;G;H;I"
                    },
                   "Lib_Macros":
                    {"Name":"Lib_Macros",
                     "Filename":"\\csv\\Lib_Macros.csv",
                     "Fieldnames": "A;B;C;D;E;F;G;H;I;J;K;L;M;N;O;P;Q;R;S;T;U;V;W;X;Y;Z;AA;AB;AC;AD;AE;AF;AG;AH;AI;AJ;AK;AL;AM;AN;AO;AP;AQ;AR;AS"
                    },
                   "Libraries":
                    {"Name":"Libraries",
                     "Filename":"\\csv\\Libraries.csv",
                     "Fieldnames": "A;B;C;D;E;F;G;H;I;J",
                     "Formating" : {"ProtectedCells"  : ((0,"*"),(1,"*"),(2,"*"),(3,"*"),(4,"*"),(5,"*"),(6,"*"),(7,"*")),
                                    "FontColor"       : { "1": {
                                                                "font"     : ("Arial",10),
                                                                "fg"       : "#FFFF00",
                                                                "bg"       : "#0000FF",
                                                                "Cells"    : ((6,"*"),)
                                                                },
                                                        "default": {
                                                               "font"     : ("Arial",10),
                                                               "fg"       : "#000000",
                                                               "bg"       : "#FFFFFF",
                                                               "Cells"    : (("*","*"),)
                                                               }
                                                        }
                                    }
                    },                   
                   "Par_Description":
                    {"Name":"Par_Description",
                     "Filename":"\\csv\\Par_Description.csv",
                      "Fieldnames": "A;B;C;D;E;F;G;H;I;J;K"
                    },                  
                   "Platform_Parameters":
                    {"Name":"Platform_Parameters",
                     "Filename":"\\csv\\Platform_Parameters.csv",
                     "Fieldnames": "A;B;C;D;E"
                    }
                }

        self.Name = "PyProgramWorkbook"
        Workbooks.append(self)
        if frame != None:
            self.Path = path
            #self.tablemodel = init_tablemodel_DCC() #get_globaltabelmodel()
            
            #self.sheets = [CWorksheet("DCC",tablemodel=self.tablemodel,frame=frame)]
            
            style = ttk.Style(frame)
            style.configure('downtab.TNotebook', tabposition='sw')
            self.container = ttk.Notebook(frame, style="downtab.TNotebook")

            self.container.grid(row=0,column=0,columnspan=2,sticky="nesw")
            self.tabdict = dict()
            self.tabdict_frames = dict()
            self.sheets = list()
        
            for sheetname in sheetdict.keys():
                sheetname_prop = sheetdict.get(sheetname)
                formating_dict = sheetname_prop.get("Formating",None)
                fieldnames = sheetname_prop.get("Fieldnames",None)
                if type(fieldnames) == str:
                    fieldnames = fieldnames.split(";")
                tabframe = ttk.Frame(self.container,relief="ridge", borderwidth=1)
                self.sheets.append(CWorksheet(sheetname,filepathname=path + sheetname_prop["Filename"],frame=tabframe,fieldnames=fieldnames,formating_dict=formating_dict))
                
                self.container.add(tabframe, text=sheetname)
            
            #self.container.bind("<<NotebookTabChanged>>",self.TabChanged)
               
        else:
            self.Path = path
            self.tablemodel= None
            self.sheets = None
        
    def Sheets(self,name):
        if self.sheets != None:
            for sheet in self.sheets:
                if sheet.Name == name:
                    return sheet
        return None
    
    def Worksheets(self,name):
        if self.sheets != None:
            for sheet in self.sheets:
                if sheet.Name == name:
                    return sheet
        return None    
        

class CWorksheet:
    def __init__(self,Name,tablemodel=None,filepathname=None,frame=None,fieldnames=None,formating_dict=None):
        
        self.width = 600
        self.height = 540
        self.ProtectContents = False
        self.Name = Name
        if tablemodel:
            self.tablemodel = tablemodel
            self.table = TableCanvas(frame, model=tablemodel,width=self.width,height=self.height,scrollregion=(0,0,self.width,self.height))
        else:
            if filepathname:
                self.tablemodel = TableModel()
                self.table = TableCanvas(frame, model=tablemodel,width=self.width,height=self.height,scrollregion=(0,0,self.width,self.height))
                self.table.importCSV(filename=filepathname, sep=';',fieldnames=fieldnames)
                self.tablemodel = self.table.getModel()
            else:
                return
        if formating_dict:
            self.tablemodel.nodisplay = formating_dict.get("HideCells",[])
            self.tablemodel.protected_cells = formating_dict.get("ProtectedCells",[])
            self.tablemodel.format_cells = formating_dict.get("FontColor",{})
        self.table.show()
        self.LastUsedRow = self.tablemodel.getRowCount()
        self.LastUsedColumn = self.tablemodel.getColumnCount()
        self.UsedRange = CRange((0,0) , (self.LastUsedRow,self.LastUsedColumn),ws=self)
        self.MaxRange  = CRange((0,0) , (self.LastUsedRow,self.LastUsedColumn),ws=self)
        self.Rectangles = CRectangles(0)
        self.AutoFilterMode = False
        self.searchcache = {}
        self.Shapes = []
        self.CellDict = CCellDict()
        self.End = self.LastUsedColumn
       
           
    def Cells(self,row, col):
        #print("Cells", self.Name, row, col)
        if row <=self.tablemodel.getRowCount() and col <=self.tablemodel.getColumnCount():
            cell = CCell(self.tablemodel.getCellRecord(row-1, col-1))
        else:
            cell=CCell("")
        cell.set_tablemodel(self.tablemodel)
        cell.set_row(row)
        cell.set_column(col)
        cell.set_parent(self)
        return cell
    
    def Range(self,cell1,cell2):
        #print("Range,cell1,cell2,ws=self")
        return CRange((cell1.Row,cell1.Column),(cell2.Row,cell2.Column),ws=self)
        
        r=[]
        for row in range(cell1.Row,cell2.Row+1):
            for col in range(cell1.Column,cell2.Column+1):
                r.append(self.Cells(row, col))
        return r
    
    def Unprotect(self):
        self.ProtectContents = False
        
    def Protect(self, DrawingObjects=True, Contents=True, Scenarios=True, AllowFormattingCells=True, AllowInsertingHyperlinks=True):
        self.ProtectContents = True
        return
        
    def Columns(col:int):
        #print("Columns",col)
        return None
    
    def EnableDisableAllButtons(self,value):
        return
    
    def SetChangedCallback(self,callback):
        self.wschanged_callback = callback
        
    def SetSelectedCallback(self,callback):
        self.wsselected_callback = callback
        
    def EventWSchanged(self, changedcell):
        if self.wschanged_callback and Application.EnableEvents:
            self.wschanged_callback(changedcell)
            
    def EventWSselected(self, selectedcell):
        if self.wsselected_callback and Application.EnableEvents:
            self.wsselected_callback(selectedcell)
            
    def Redraw_table(self):
        self.table.redraw()
        
    def set_value_in_cell(self,row,column,newval):
        colname = self.tablemodel.getColumnName(column-1)
        #coltype = tablemodel.columntypes[colname]
        name = self.tablemodel.getRecName(row-1)
        if colname in self.tablemodel.data[name]:
            if self.tablemodel.data[name][colname] != newval:
                self.tablemodel.data[name][colname] = newval
                if Application.EnableEvents:
                    print("Workbook changed")
                    ActiveSheet.EventWSchanged(self)
        
    def create_search_colcache(self,searchcol):
        colcontent = self.tablemodel.getColCells(searchcol-1)
        row = 1
        self.searchcache[searchcol]={}
        searchcol_dict = self.searchcache[searchcol]
        for col in colcontent:
            searchcol_dict[col]=row
            row=row+1
        return
        
    def find_in_col_ret_col_val(self,searchtext, searchcol, resultcol,cache=True):
        
        if cache:
            colcache = self.searchcache.get(searchcol,None)
            if not colcache:
                self.create_search_colcache(searchcol)
                colcache = self.searchcache.get(searchcol,None)
            res_row = colcache.get(searchtext,None)
            if not res_row:
                return None # searchtext not found
            else:
                return self.tablemodel.getCellRecord(res_row-1, resultcol-1)
        else:
            pass
        return None
                
    def find_in_col_set_col_val(self,searchtext, searchcol, setcol,setval,cache=False):
        
        if cache:
            colcache = self.searchcache.get(searchcol,None)
            if not colcache:
                self.create_search_colcache(searchcol)
                colcache = self.searchcache.get(searchcol,None)
            res_row = colcache.get(searchtext,None)
            if not res_row:
                return None # searchtext not found
            else:
                self.set_value_in_cell(res_row, setcol, setval)
                return True
        else:
            pass
        return None
    
    def Activate(self):
        return
    
    def Select(self):
        return
        
        

class CRange:
    def __init__(self,t1,t2,ws=None):
        #print("Range:",t1,t2)
        self.range_list=[]
        self.Rows=[]
        self.Columns = []
        self.Cells = []
        for col in range(t1[1],t2[1]+1):
            self.Columns.append(CColumn(col))
            for row in range(t1[0],t2[0]+1):
                self.range_list.append(ws.Cells(row, col))
        for row in range(t1[0],t2[0]+1):
            self.Rows.append(CRow(row))        
        self.CountLarge = len(self.range_list)
        self.Cells = self.range_list
        self.Font = CFont("Arial",10)
                
    def Find(self,What="", LookIn=xlFormulas, LookAt= xlWhole, SearchOrder= xlByRows, SearchDirection= xlNext, MatchCase= True, SearchFormat= False):
        for cell in self.range_list:
            if cell.Value == What:
                return cell
            
    def Offset(self, offset_row, offset_col):
        for cell in self.range_list:
            cell.offset(offset_row,offset_col)
        return self
    
    def Activate(self):
        for cell in self.range_list:
            cell.Activate()        
            
class CRectangles:
    def __init__(self,count):
        self.list=[]
        self.Count=0

    #def __init__(self,rowrange,colrange):
    #    self.Rows=list()
    #    for i in range(rowrange[0],rowrange[1]):
    #        self.Rows.append(CRow(i))
    #    self.Columns=list()
    #    for i in range(colrange[0],colrange[1]):
    #        self.Columns.append(CRow(i))
class CSelection:
    def __init__(self,cell):
        self.EntireRow = CEntireRow(cell.Row)
        self.EntireColumn = CEntireColumn(cell.Column)
        self.Characters = ""
        

class CRow:
    def __init__(self,rownumber):
        self.Row = rownumber
        self.EntireRow = CEntireRow(rownumber)
        
class CEntireRow:
    def __init__(self,rownumber):
        self.Rownumber = rownumber
        self.Hidden = False

        
class CColumn:
    def __init__(self,colnumber):
        self.Column=colnumber
        self.EntireColumn = CEntireColumn(colnumber)
        
class CEntireColumn:
    def __init__(self,colnumber):
        self.Columnnumber = colnumber
        self.Hidden = False
    
    
class CCell(str):
    def __init__(self,value,tablemodel=None):
        str.__init__(value)
        self.Orientation = 0
        self.Row = -1
        self.Column = -1
        self.Value = value
        self.Formula = ""
        self.tablemodel = tablemodel
        self.CountLarge = 1
        self.Height = 20
        self.Width = 30
        self.Parent = None
        self.Top = 0
        self.Left = 0
        self.Address = (self.Row,self.Column)
        self.Comment = None
        self.Font = CFont("Arial",10)
        self.HorizontalAlignment = xlCenter
        self.Text = value
        
    def get_value(self):
        if self.Row != -1:
            if self.tablemodel == None:
                self.tablemodel = ActiveSheet.tablemodel
            max_cols = self.tablemodel.getColumnCount()
            if self.Column < max_cols:
                value = self.tablemodel.getValueAt(self.Row-1,self.Column-1)
            else:
                value = ""
        return value # self.__value
    
    def set_row(self,row):
        self.Row = row
        self.Address = (self.Row,self.Column)
        
    def set_column(self,column):
        self.Column=column
        self.Address = (self.Row,self.Column)
    
    def set_value(self, newval):
        if self.Row != -1:
            if self.tablemodel == None:
                self.tablemodel = ActiveSheet.tablemodel
            colname = self.tablemodel.getColumnName(self.Column-1)
            #coltype = tablemodel.columntypes[colname]
            name = self.tablemodel.getRecName(self.Row-1)
            if colname in self.tablemodel.data[name]:
                if self.tablemodel.data[name][colname] != newval:
                    self.tablemodel.data[name][colname] = newval
                    if type(newval) != str:
                        print("Type not str",newval)
                    if Application.EnableEvents:
                        print("Workbook changed")
                        ActiveSheet.EventWSchanged(self)
                        #M20.Global_Worksheet_Change(self)
                
    def set_tablemodel(self,tablemodel):
        self.tablemodel = tablemodel
    
    def set_parent(self,parent):
        self.Parent = parent
        
    def check_if_empty_row(self):
        if self.tablemodel == None:
            self.tablemodel = ActiveSheet.tablemodel        
        cols = self.tablemodel.getColumnCount()
        r=self.Row-1
        for c in range(0,cols):
            #absr = self.get_AbsoluteRow(r)
            val = self.tablemodel.getValueAt(r,c)
            if val != None and val != '':
                return False
        return True
    
    def Offset(self, offset_row, offset_col):
        self.Row    = self.Row + offset_row
        self.Column = self.Column + offset_col
        ActiveSheet.table.gotoCell(self.Row-1,self.Column-1)
        return ActiveCell()
    
    def Select(self):
        ActiveSheet.table.gotoCell(self.Row-1,self.Column-1)
        ActiveSheet.EventWSselected(self)
        return
    
    def AutoFilter(self):
        print("Autofilter")
        
    def Activate(self):
        self.Parent.table.gotoCell(self.Row,self.Column)
    
    Value = property(get_value, set_value, doc='value of CCell')
    Text = property(get_value, set_value, doc='value of CCell')
    
class CCellDict():
    def __init__ (self):
        data = {}
        
    def __getitem__(self, k):
        #print("Getitem",k)
        return Cells(k[0],k[1])
        
    def __setitem__(self,k,value):
        print("Setitem",k, value)
        ccell = Cells(k[0],k[1])
        ccell.set_value(value)
        
        
class CWorksheetFunction:
    def __init__(self):
        pass
    def RoundUp(v1,v2):
        #print("RoundUp")
        if v2 == 0:
            if v1 == int(v1):
                return v1
            else:
                return int(v1 + 1)
        else:
            return "Error"
        
class CApplication:
    def __init__(self):
        self.StatusBar = ""
        self.EnableEvents = True
        self.ScreenUpdating = True
        self.Top = 0
        self.Left = 0
        self.Width = 1000 #*HL
        self.Height = 500 #*HL
        self.Version = "15"
        
    def OnTime(self,time,cmd):
        print("Application OnTime:",time,cmd)
        return #*HL
    
    def RoundUp(v1,v2):
        #print("RoundUp")
        if v2 == 0:
            if v1 == int(v1):
                return v1
            else:
                return int(v1 + 1)
        else:
            return "Error"
        
class CFont:
    def __init__(self,name,size):
        self.Name = name
        self.Size = size
        
class SoundLines:
    def __init__(self):
        self.dict = {}
        self.Count = 0
        self.keys = []
    
    def Exists(self,Channel:int):
        pass
    
    def Add(Channel, Pin_playerClass):
        pass
    

# global variables

ActiveWorkbook:CWorkbook = None

ActiveSheet:CWorksheet = None

WorksheetFunction = CWorksheetFunction()

Application:CApplication = CApplication()

CellDict:CCellDict = CCellDict()

Selection = CSelection(CCell(""))

Workbooks = []


Now = 0
