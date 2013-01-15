#!/usr/bin/python -tt
# -*- coding: utf-8 -*-
###########################################################
# Copyright 2012 Florian PERVES
# 
# This file is part of PYNBM.
# 
# PYNBM is free software: you can redistribute it and/or
# modify it under the terms of the GNU General Public License
# as published by the Free Software Foundation, either version
# 3 of the License, or (at your option) any later version.
# 
# PYNBM is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty
# of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See
# the GNU General Public License for more details.
# 
# You should have received a copy of the GNU General Public
# License along with PYNBM. If not, see
# http://www.gnu.org/licenses/.
###########################################################


###########################################################
# 2012 - Florian PERVES - florian.perves@gmail.com 
###########################################################
#
#Changelog
#1.1.1 : - launch custom tool without console window on MS Windows
#1.1   : - launch custom tool on several selected items
#        - set the result of the custom tool command in the grid
#        - load last opened file on startup (option)
#        - resize column width on label double click
#        - resize cells width and height with a new Icon
#        - turn passwd cell in white on import if 'hide passwd' mode is activated
#        - multiple lines in a cell (ALT+ENTER to enter a new line)
#        - reset stat turns all items in black
#1.0.5 : - turn cell text to black when importing or re-ordering.
#        - embbed ico in exe
#        - application icon on windows 7
#        - reset stats function
#1.0.4 : - update window title on file import/export.
#        - fixed v1.0.3 installer incompatibility with Windows XP
#1.0.3 : - use MS WMI to call ping function on windows
#        - optimization of the stop/start methods for polling threads
#        - bring portable option (-p or -portable) and thus use current path to store
#          configuration files, output logs, etc.
#        - add PYNBM entry in MS Windows add/remove programs utility
#        - add an exit dialog prompting to save the grid
#1.0.2 : - bug fix http/https access when URL field is filled
#1.0.1 : - bug fix in sub menus in context menu not working on Windows
#1.0 :   - sub menus in context menu        
#        - auto-update
#0.2.8 : - ignore CSV columns that exceed Grid size (used to crash import)
#        - find is now case-insensitive
#        - removed trailing newline char in copy method when only one cell is selected
#0.2.7 : - search algorithm improvements
#        - Start first polling before waiting interval
#        - creation of autonomous RDP and AES Classes in separate files
#0.2.6 : - Number of line configurable from settings view
#        - Custom Columns configurable from settings view
#        - Ascending/Descending sort of column contents by right-clicking column header
#        - Columns can be reordered by drag and drop
#        - Select all cell with CTRL-a
#        - Moved all configuration files to Application Data\PYNBM directory (Windows), ~/Library/Application Support/PYNBM' (Mac OSX).
#        - Polling threads enhancements : loop interval x sleep 1 sec, and check 
#          polling is still enabled at each loop
#        - stderr redirected to logfile in Application Data dir for windows
#        - no overwrite of Extensions.conf, PYNBM.ini
#0.2.5 : - check file existence before export, and prompt before overwrite,
#        - fix grid locked bug, which used to allow deletion of cell contents
#          even when locked
#0.2.4 : - lock the grid

# Python standard libraries
import ConfigParser
import time
import threading
import re
import sys
import os
import webbrowser
import subprocess
from decimal import Decimal
#import smtplib
import logging
import shutil
#import operator
import urllib2




# WxPython libraries
import wx
import wx.grid
import wx.lib.newevent
from wx.lib.wordwrap import wordwrap
import wx.lib.buttons as buttons

#PYNBM AES specific library
import pynbm_aes
from pynbm_CustomContextItem import pynbm_CustomContextItem
from pynbm_refresh import pynbm_refresh
from pynbm_updateChecker import pynbm_updateChecker

# global variables
Name = "PYNBM"
Version = "1.1.1"
Release_Date = "Dec 31 2012"
downloadURL = 'http://www.pynbm.org/download'


ERRORLOG = 'PYNBM_errors.log'
LOG = 'PYNBM.log'

ID_IMPORTCSV = wx.NewId()
ID_EXPORTCSV = wx.NewId()
ID_TOGGLEPOLLING = wx.NewId()
ID_TOGGLEHIDEPASSWD = wx.NewId()
ID_SETTINGS = wx.NewId()
ID_GRIDLOCK = wx.NewId()
ID_RESETSTAT = wx.NewId()
ID_RESIZE = wx.NewId()

ID_ABOUT = wx.NewId()
ID_SEARCH = wx.NewId()
ID_SEARCHNEXT = wx.NewId()
ID_COPY = wx.NewId()
ID_PASTE = wx.NewId()
ID_SELECTALL = wx.NewId()
ID_HTTP = wx.NewId()
ID_HTTPS = wx.NewId()

EventID = {'ID_IMPORTCSV':ID_IMPORTCSV, 
            'ID_EXPORTCSV':ID_EXPORTCSV, 
            'ID_TOGGLEPOLLING':ID_TOGGLEPOLLING, 
            'ID_SETTINGS':ID_SETTINGS, 
            'ID_GRIDLOCK':ID_GRIDLOCK, 
            'ID_RESETSTAT':ID_RESETSTAT,
            'ID_RESIZE':ID_RESIZE,
            'ID_ABOUT':ID_ABOUT,
            'ID_SEARCH':ID_SEARCH,
            'ID_COPY':ID_COPY,
            'ID_PASTE':ID_PASTE,
            'ID_SELECTALL':ID_SELECTALL,
            'ID_HTTP':ID_HTTP,
            'ID_HTTPS':ID_HTTPS}


#threadlist = []

ostype = sys.platform

EXTENSIONS_FILENAME = 'Extensions.conf'
INI_FILENAME = 'PYNBM.ini'

interval = None
timeout = None
numberOfRefreshThreads = None
logtofile = None
logfile = None
checkForUpdate = None
customColumns = []
numberOfLines = None
logger = logging.getLogger('PYNBM')
logger.setLevel(logging.DEBUG)
fh=None #logger file handler

if (ostype == "win32"):
    #PYNBM specific rdp library
    import pynbm_rdp
    #PYNBM specific wmiping library
    import pynbm_wmiping
    appDataPath=os.environ['APPDATA']
    tmpPath=os.environ['TMP']
    EXTENSIONS_FILENAME = appDataPath + '\PYNBM\Extensions.conf'
    INI_FILENAME = appDataPath + '\PYNBM\PYNBM.ini'
    checkVersionURL = 'http://www.pynbm.org/data/documents/current_windows_version.html'
    if (len(sys.argv) > 1):
        if sys.argv[1] in ['-p', '-portable']:
            sys.stderr = open('PYNBM_stderr.log', 'w')
            sys.stdout = open('PYNBM_stdout.log', 'w')
            EXTENSIONS_FILENAME = '.\Extensions.conf'
            INI_FILENAME = '.\PYNBM.ini'
        elif sys.argv[1] in ['-d', '-debug']:
            sys.stderr = sys.__stderr__
            sys.stdout = sys.__stdout__
    else:
       sys.stderr = open(appDataPath + '\PYNBM\PYNBM_stderr.log', 'w')
       sys.stdout = open(appDataPath + '\PYNBM\PYNBM_stdout.log', 'w')    
       
elif (ostype == "darwin"):
    appDataPath=os.path.expanduser('~') + '/Library/Application Support/PYNBM/'
    if not os.path.exists(appDataPath):
        os.makedirs(appDataPath)
    checkVersionURL = 'http://www.pynbm.org/data/documents/current_mac_version.html'
    if (len(sys.argv) > 1):
        if sys.argv[1] in ['-p', '-portable']:
            sys.stderr = open('PYNBM_stderr.log', 'w')
            sys.stdout = open('PYNBM_stdout.log', 'w')
            EXTENSIONS_FILENAME = 'Extensions.conf'
            INI_FILENAME = 'PYNBM.ini'
        elif sys.argv[1] in ['-d', '-debug']:
            sys.stderr = sys.__stderr__
            sys.stdout = sys.__stdout__
    else:
       sys.stderr = open(appDataPath + 'PYNBM_stderr.log', 'w')
       sys.stdout = open(appDataPath + 'PYNBM_stdout.log', 'w')    
    EXTENSIONS_FILENAME = appDataPath + EXTENSIONS_FILENAME
    INI_FILENAME = appDataPath + INI_FILENAME

else :
    appDataPath=os.path.expanduser('~') + '/.PYNBM/'
    if not os.path.exists(appDataPath):
        os.makedirs(appDataPath)
    checkVersionURL = 'http://www.pynbm.org/data/documents/current_linux_version.html'
    if (len(sys.argv) > 1):
        if sys.argv[1] in ['-p', '-portable']:
            sys.stderr = open('PYNBM_stderr.log', 'w')
            sys.stdout = open('PYNBM_stdout.log', 'w')
            EXTENSIONS_FILENAME = 'Extensions.conf'
            INI_FILENAME = 'PYNBM.ini'
        elif sys.argv[1] in ['-d', '-debug']:
            sys.stderr = sys.__stderr__
            sys.stdout = sys.__stdout__
    else:
       sys.stderr = open(appDataPath + 'PYNBM_stderr.log', 'w')
       sys.stdout = open(appDataPath + 'PYNBM_stdout.log', 'w')    
    EXTENSIONS_FILENAME = appDataPath + EXTENSIONS_FILENAME
    INI_FILENAME = appDataPath + INI_FILENAME

(UpdateRTTEvent, EVT_UPDATE_RTT) = wx.lib.newevent.NewEvent()
(UpdateColEvent, EVT_UPDATE_CELL) = wx.lib.newevent.NewEvent()
(ThreadEndEvent, EVT_THREAD_END) = wx.lib.newevent.NewEvent()
(NewVersion, EVT_NEWVERSION) = wx.lib.newevent.NewEvent()
(ExternalToolStatusEvent, EVT_EXTERNALTOOLSTATUS) = wx.lib.newevent.NewEvent()

class IniFile():
    def __init__(self,filename):
        self.filename = filename
        
    def load_from_file(self):
        self.config = ConfigParser.SafeConfigParser()
        global interval
        global timeout
        global numberOfRefreshThreads
        global logtofile
        global checkForUpdate
        global logfile
        global lastOpenedFile
        global customColumns
        global numberOfLines
        global loadLastOpenedFileOnStartup
        try:
            self.config.read(self.filename)
        except:
            interval = 10
            timeout = 3
            numberOfRefreshThreads = 10
            logtofile = 0
            checkForUpdate = True
            logfile = None
            lastOpenedFile = None
            customColumns = []
            numberOfLines = 1000
            loadLastOpenedFileOnStartup = False
        try:
            interval = self.config.getint('Poller', 'interval')
        except:
            interval = 10       
        try:
            timeout = self.config.getint('Poller', 'timeout')
        except:
            timeout = 3        
        try:
            numberOfRefreshThreads = self.config.getint('Poller', 'threads')
        except:
            numberOfRefreshThreads = 10
        try:
            logtofile = self.config.getboolean('Poller', 'logtofile')
        except:
            logtofile = False 
        try:
            logfile = self.config.get('Poller', 'logfile')
        except:
            logfile = None
        try:
            lastOpenedFile = self.config.get('PYNBM', 'lastOpenedFile')
        except:
            lastOpenedFile = None
        try:
            checkForUpdate = self.config.getboolean('PYNBM', 'checkForUpdate')
        except:
            checkForUpdate = True 
        try:
            loadLastOpenedFileOnStartup = self.config.getboolean('PYNBM', 'loadLastOpenedFileOnStartup')
        except:
            loadLastOpenedFileOnStartup = False     
        try:
            numberOfLines = self.config.getint('Grid', 'lines')
        except:
            numberOfLines = 1000
        try:
            customColumns = self.config.get('Grid', 'customColumns')
            if (customColumns != ""):
                customColumns = customColumns.split(";")
        except:
            customColumns = []
        self.save_to_file()

    def save_to_file(self):
        self.config = ConfigParser.SafeConfigParser()
        self.config.add_section('PYNBM')
        self.config.set('PYNBM', 'checkForUpdate', str(checkForUpdate))
        self.config.set('PYNBM', 'loadLastOpenedFileOnStartup', str(loadLastOpenedFileOnStartup))
        self.config.set('PYNBM', 'lastOpenedFile', str(lastOpenedFile))        
        self.config.add_section('Poller')
        global interval
        global timeout
        self.config.set('Poller', 'interval', str(interval))
        self.config.set('Poller', 'timeout', str(timeout))
        self.config.set('Poller', 'threads', str(numberOfRefreshThreads))
        self.config.set('Poller', 'logtofile', str(logtofile))
        self.config.set('Poller', 'logfile', str(logfile))
        self.config.add_section('Grid')
        self.config.set('Grid', 'customColumns', str(";".join(customColumns)))     
        self.config.set('Grid', 'lines', str(numberOfLines))      
        d = os.path.dirname(self.filename)
        if (ostype == "win32" and not os.path.exists(d)):
            os.makedirs(d)
        with open(self.filename, 'wb') as configfile:
            self.config.write(configfile)
        global fh
        if fh != None:
            logger.removeHandler(fh)
        if (str(logtofile) == "True"):
            fh = logging.FileHandler(str(logfile))
            fh.setLevel(logging.DEBUG)
            formatter = logging.Formatter('%(asctime)s - %(message)s')
            fh.setFormatter(formatter)
            logger.addHandler(fh)

class Grid(wx.grid.Grid):
    def __init__(self, parent):
        """
        initialization of the grid, definition of column names, mouse 
        right-click events binding
        """
        wx.grid.Grid.__init__(self, parent, -1)
        self.colnames = ['host','alias','login','password','url','status',
                         'RTT','rate']
        self.mainwindow = parent
        self.colnames.extend(customColumns)
        self.CreateGrid(numberOfLines, len(self.colnames))
        self.editor = wx.grid.GridCellAutoWrapStringEditor() 
        self.SetDefaultEditor(self.editor) 
        self.Edition="enabled"
        colid = 0
        for colname in self.colnames:
            self.SetColLabelValue(colid,colname)
            colid += 1
        self.Bind(wx.grid.EVT_GRID_CELL_RIGHT_CLICK, self.on_cell_right_click)
        self.Bind(wx.grid.EVT_GRID_LABEL_LEFT_DCLICK, self.resize_columns)
        self.Bind(wx.grid.EVT_GRID_LABEL_RIGHT_CLICK, self.on_label_right_click)
        self.EnableDragColMove()
        # invokes on_key method when a key is pressed
        wx.EVT_KEY_DOWN(self, self.on_key)
        self.patternToSearch=""

    def corners_to_cells(self,top_lefts, bottom_rights):
        """
        Take lists of top left and bottom right corners, and
        return a list of all the cells in that range
        """
        cells = []
        for top_left, bottom_right in zip(top_lefts, bottom_rights):

            rows_start = top_left[0]
            rows_end = bottom_right[0]

            cols_start = top_left[1]
            cols_end = bottom_right[1]

            rows = range(rows_start, rows_end+1)
            cols = range(cols_start, cols_end+1)

            cells.extend([(row, col)
                for row in rows
                for col in cols])
    
        return cells
    
    def get_selected_rows(self):
        selection=self.get_selected_cells()
        rows=[]
        for (row,column) in selection:
            rows.append(row)
        
        #remove duplicates:
        rows = list(set(rows))
        return rows

    def resize(self):
        for colid in range(len(self.colnames)):
            self.AutoSizeColumn(colid,False)
        for rowid in range(self.mainwindow.get_last_line_number()):
            self.AutoSizeRow(rowid,False)
        self.ForceRefresh()

    def get_selected_cells(self):
        """
        Return the selected cells in the grid as a list of
        (row, col) pairs.
        We need to take care of three possibilities:
        1. Multiple cells were click-selected (GetSelectedCells)
        2. Multiple cells were drag selected (GetSelectionBlock…)
        3. A single cell only is selected (CursorRow/Col)
        """

        top_left = self.GetSelectionBlockTopLeft()
    
        if top_left:
            bottom_right = self.GetSelectionBlockBottomRight()
            return self.corners_to_cells(top_left, bottom_right)

        selection = self.GetSelectedCells()

        if not selection:
            row = self.GetGridCursorRow()
            col = self.GetGridCursorCol()
            return [(row, col)]

        return selection
        
    def xy_to_cell(self, x, y):
        """
        Given a position relative to upper left corner of grid, return 
        the cell row and col of that position as a tuple (row, col).
        The origin of the input position is the upper edge of the column label
        row, and the left edge of the row label column.
        """
        # we need to compute the 'logical' position. In this coordinate system,
        # the origin is the upper left corner of cell 0:0.
        logicalPosX, logicalPosY = self.CalcUnscrolledPosition(x, y)
        # remove the label height/width from the logical pos
        # (so x=0, y=0 is upper left corner of cell at row=0, col=0)
        logicalPosX -= self.GetRowLabelSize()
        logicalPosY -= self.GetColLabelSize()
        row = self.YToRow(logicalPosY)
        col = self.XToCol(logicalPosX)
        return (row, col)
        
    def on_cell_right_click(self, evt):
        """ if right mouse button is clicked out of the currently selected 
        cell, move selection to the cell that is under the mouse at 
        right-click time
        """
        pos = evt.GetPosition()
        row, col = self.xy_to_cell(pos.x, pos.y)
        if row >= 0 and col >= 0:
            self.SetGridCursor(row, col)
        self.show_context_menu()


    def on_label_right_click(self,evt):
        col = evt.GetCol()
        x = self.GetColSize(col)/2
        menu = wx.Menu()
        ascSortID = wx.NewId()
        dscSortID = wx.NewId()
        xo, yo = evt.GetPosition()
#        self.SelectCol(col)
#        cols = self.GetSelectedCols()
        self.Refresh()
        menu.Append(ascSortID, "Sort Ascending")
        menu.Append(dscSortID, "Sort Descending")
        
        def ascending(evt):
            columnSort(reverse=False)
        
        def descending(evt):
            columnSort(reverse=True)
        
        def columnSort(reverse, self=self,col=col):
            res = self.return_table()
            col = col
            reverse = reverse        
            if reverse==True:
                sortedgrid = sorted(res,key=lambda x: (x[col] != "", x[col]),reverse=True)
            else:
                sortedgrid = sorted(res,key=lambda x: (x[col] == "", x[col]))
            i = 0        
            for row in sortedgrid:
                line = []
                j = 0            
                for cell in row:
                    self.SetCellValue(i, j, cell)
                    #print str(i) + ";" + str(j)
                    self.SetCellTextColour(i, j, wx.BLACK)
                    j=j+1
                i=i+1


        self.Bind(wx.EVT_MENU, ascending, id=ascSortID)        
        self.Bind(wx.EVT_MENU, descending, id=dscSortID)        
        self.PopupMenu(menu)
        menu.Destroy()
        return




    def return_table(self):
        gridArray=[]
        
        for r in range(numberOfLines):
            gridArray.append([])
            for c in range(len(self.colnames)):
                val = self.GetCellValue(r,c)
                gridArray[r].append(val)
        return gridArray                

    def resize_columns(self,evt):
        self.AutoSizeColumn(evt.GetCol(),False)

        
    def import_extensions(self,menu):
        """
        read extensions definitions from EXTENSION_FILENAME, and then create
        CustomContextItem objets fore each definition
        """
        try:
            f = open(EXTENSIONS_FILENAME, 'r')
        except IOError:
            shutil.copy2('Extensions.defaults', EXTENSIONS_FILENAME)
            self.import_extensions(menu)
        else:
            self.mainMenu = menu
            self.currentMenu = menu
            self.inSubMenu = False
            self.inSubMenuName = ""          
            for line in f:                    
                if (re.search("^[^#]",line)):
                    line=line.rstrip()
                    if (re.search("^\[.+\]",line)):
                        contextMenu = pynbm_CustomContextItem(self.mainwindow,UpdateColEvent,ExternalToolStatusEvent)
                        self.currentMenu = contextMenu.create_item("subMenu",self.currentMenu,self,self)
                        self.inSubMenuName = line
                        self.inSubMenu = True
                
                    elif (re.search("^$",line) and (self.inSubMenu is True)):
                        contextMenu = pynbm_CustomContextItem(self.mainwindow,UpdateColEvent,ExternalToolStatusEvent)                  
                        res = contextMenu.create_item("subMenuEnd",self.currentMenu,self.mainMenu,self,self.inSubMenuName)
                        self.currentMenu = self.mainMenu
                        self.inSubMenu = False
                        self.inSubMenuName = ""   
                
                    elif (re.search("^$",line) and (self.inSubMenu is False)):
                        pass
                
                    else:
                        cells = line.split('=',1)
                        funcname = cells[0].rstrip('\n')
                        commandline = ""
                        if (len(cells) == 2):
                            commandline = cells[1].rstrip('\n')
                        contextMenu = pynbm_CustomContextItem(self.mainwindow,UpdateColEvent,ExternalToolStatusEvent)
                        res = contextMenu.create_item(funcname,commandline,self.currentMenu,self)
            f.close()

    def show_context_menu(self):
        """
        Create and display a popup menu on right-click event
        """
        menu = wx.Menu()
        contexthttp = wx.MenuItem(menu, wx.NewId(), 'http')
        menu.AppendItem(contexthttp)
        menu.Bind(wx.EVT_MENU, self.http, contexthttp)

        contexthttps = wx.MenuItem(menu, wx.NewId(), 'https')
        menu.AppendItem(contexthttps)
        menu.Bind(wx.EVT_MENU, self.https, contexthttps)

        if (ostype == "win32"):
            contextrdp = wx.MenuItem(menu, wx.NewId(), 'RDP')
            menu.AppendItem(contextrdp)
            menu.Bind(wx.EVT_MENU, self.rdp, contextrdp)

        self.import_extensions(menu)

        menu.AppendSeparator()

        contextcopy = wx.MenuItem(menu, wx.NewId(), u'Copy')
        menu.AppendItem(contextcopy)
        menu.Bind(wx.EVT_MENU, self.copy, contextcopy)
        
        contextpaste = wx.MenuItem(menu, wx.NewId(), u'Paste')
        menu.AppendItem(contextpaste)
        menu.Bind(wx.EVT_MENU, self.paste, contextpaste)
        
        contextinsert = wx.MenuItem(menu, wx.NewId(), u'Insert line')
        menu.AppendItem(contextinsert)
        menu.Bind(wx.EVT_MENU, self.insert_row, contextinsert)

        contextdelete = wx.MenuItem(menu, wx.NewId(), u'Delete line')
        menu.AppendItem(contextdelete)
        menu.Bind(wx.EVT_MENU, self.delete_row, contextdelete)

        self.PopupMenu(menu)
        menu.Destroy()

    def insert_row(self,evt):
        global numberOfLines
        self.InsertRows(self.GetGridCursorRow(),1)
        numberOfLines = numberOfLines +1

    def delete_row(self,evt):
        global numberOfLines
        self.DeleteRows(self.GetGridCursorRow(),1)
        numberOfLines = numberOfLines -1
    
        

    def unlock_edit(self):
        self.Edition="enabled"              
        self.EnableEditing(True)

    def lock_edit(self):
        self.Edition="disabled"
        self.EnableEditing(False)
    
    def rdp(self,evt):
        for row in (self.get_selected_rows()):
            host=self.GetCellValue(row,0)
            rdpfilename=os.path.join(tmpPath,host.split(":")[0]+".rdp")
            if (len(self.GetCellValue(row,2).split('\\')) > 1):
                domain = self.GetCellValue(row,2).split('\\')[0]
                login = self.GetCellValue(row,2).split("\\")[1]
            else:
                domain = ""
                login = self.GetCellValue(row,2)
            password = self.GetCellValue(row,3)
            timer=5
            pynbm_rdp.RDP(host,domain,login,password,rdpfilename,timer)

    def http(self,evt):
        """
        points default web browser to http://host
        """
        for row in (self.get_selected_rows()):
            if self.GetCellValue(row,4) != "":
                host = self.GetCellValue(row,4)
            else:
                host = "http://"+self.GetCellValue(row,0)
            webbrowser.open("%s" % host, new = 0)

    def https(self,evt):
        """
        points default web browser to https://host
        """
        for row in (self.get_selected_rows()):
            if self.GetCellValue(row,4) != "":
                host = self.GetCellValue(row,4)
            else:
                host = "http://"+self.GetCellValue(row,0)
            webbrowser.open("%s" % host, new = 0)

    def on_key(self, event):
        # If Suppr is pressed
        if event.GetKeyCode() == 127:
            if self.Edition == "enabled" :
                self.delete()

        elif event.AltDown() and event.GetKeyCode() == wx.WXK_RETURN:
            self.editor.StartingKey(event) 
        else:         # Skip other Key events
            event.Skip()
            return

    def copy(self,evt=None):
        # If selection contains more than one cell
        if len(self.GetSelectionBlockBottomRight()) > 0: 
            rows = self.GetSelectionBlockBottomRight()[0][0] - self.GetSelectionBlockTopLeft()[0][0] + 1
            cols = self.GetSelectionBlockBottomRight()[0][1] - self.GetSelectionBlockTopLeft()[0][1] + 1

            # data variable contain text that must be set in the clipboard
            data = ''

            # For each cell in selected range, append the cell value in the 
            # data variable
            # Tabs '\t' for cols and '\r' for rows
            for r in range(rows):
                for c in range(cols):
                    data = data + str(self.GetCellValue(self.GetSelectionBlockTopLeft()[0][0] + r,
                        self.GetSelectionBlockTopLeft()[0][1] + c))
                    if c < cols - 1:
                        data = data + '\t'
                data = data + '\n'
        else: #If selection contains only one cell
            data = str(self.GetCellValue(self.GetGridCursorRow(),
                self.GetGridCursorCol()))
        # Create text data object
        clipboard = wx.TextDataObject()
        # Set data object value
        clipboard.SetText(data)
        # Put the data in the clipboard
        if wx.TheClipboard.Open():
            wx.TheClipboard.SetData(clipboard)
            wx.TheClipboard.Close()
        else:
            wx.MessageBox("Can't open the clipboard", "Error")

    def paste(self,evt=None):
        if self.Edition == "enabled" :
            clipboard = wx.TextDataObject()
            if wx.TheClipboard.Open():
                wx.TheClipboard.GetData(clipboard)
                wx.TheClipboard.Close()
            else:
                wx.MessageBox("Can't open the clipboard", "Error")
            data = clipboard.GetText()
            table = []
            y = -1

            # Convert text in an array of lines
            for r in data.splitlines():
                y = y +1
                x = -1
                # Convert c in an array of text separated by tab
                for c in r.split('\t'):
                    x = x +1
                    if ((self.GetGridCursorCol() + x) >= (len(self.colnames))): # si le paste essaie d'ecrire au del�  de la grille
                        continue
                    self.SetCellValue(self.GetGridCursorRow() + y,
                        self.GetGridCursorCol() + x, c)

    def delete(self):
        # If selection contains more than one cell
        if len(self.GetSelectionBlockBottomRight())>0: 
            # Number of rows and cols
            rows = self.GetSelectionBlockBottomRight()[0][0] - self.GetSelectionBlockTopLeft()[0][0] + 1
            cols = self.GetSelectionBlockBottomRight()[0][1] - self.GetSelectionBlockTopLeft()[0][1] + 1
            # Clear cells contents
            for r in range(rows):
                for c in range(cols):
                    self.SetCellValue(self.GetSelectionBlockTopLeft()[0][0] + r,
                        self.GetSelectionBlockTopLeft()[0][1] + c, '')
        else: # If selection contains only one cell
            self.SetCellValue(self.GetGridCursorRow(),self.GetGridCursorCol(),'')

    def selectAll(self):
        self.SelectAll()            
            
    def define_search(self):
        dlg = wx.TextEntryDialog(self, u"String to search for (case-insensitive) :")    
        self.MakeCellVisible(self.GetGridCursorRow(),self.GetGridCursorCol())                    
        if dlg.ShowModal() == wx.ID_OK:  
            self.patternToSearch = dlg.GetValue()
            dlg.Destroy()
            self.search()
            
    def search(self):
        if (self.patternToSearch == ""):
            dlg = wx.MessageDialog(self, 
                              "Are you trying to find an empty string...?",
                              u"Sorry",
                              wx.OK | wx.ICON_INFORMATION
            #wx.YES_NO | wx.NO_DEFAULT | wx.CANCEL | wx.ICON_INFORMATION
                                  )
            dlg.ShowModal()
            dlg.Destroy()
            return
        
        found = 0
        searchStartRow = self.GetGridCursorRow()
        searchEndRow = numberOfLines
        searchStartCol = self.GetGridCursorCol()+1
        searchEndCol = len(self.colnames)
        
        if searchStartCol == searchEndCol:
            searchStartCol = 0
            searchStartRow = searchStartRow+1
            if searchStartRow == searchEndRow+1:
                searchStartRow = 0

        def searchLoop(searchStartRow,searchEndRow,searchStartCol,searchEndCol):
        #Loop on grid cells searching for pattern.
            found = 0
            for linenumber in range(searchStartRow,searchEndRow):
                for col in range(searchStartCol,searchEndCol):
                    searchStartCol = 0
                    if (re.search(self.patternToSearch,self.GetCellValue(linenumber,col),re.IGNORECASE)):
                        self.SetGridCursor(linenumber, col)
                        self.MakeCellVisible(linenumber, col)
                        found = 1
                        break
                if (found == 1):
                    break
            if (found == 1):
                return 1
            else:
                return 0

                    
        searchres = searchLoop(searchStartRow,searchEndRow,searchStartCol,searchEndCol)
        if searchres == 0:
            searchEndRow = searchStartRow
            searchStartRow = 0
            searchStartCol = 0
            searchres = searchLoop(searchStartRow,searchEndRow,searchStartCol,searchEndCol)
           
        if (searchres == 0):
            dlg = wx.MessageDialog(self, 
                              "String " + self.patternToSearch + u" not Found ! ",
                              u"Not found",wx.OK | wx.ICON_INFORMATION)
            dlg.ShowModal()
            dlg.Destroy()
        
class MainWindow(wx.Frame):
    def __init__(self, titre):
        """
        build the application main window by defining the menus, buttons,
        status bar, event handlers
        """
        self.titre = titre
        frame = wx.Frame.__init__(self, None, -1, title = self.titre, size = (800, 600))
        self.gridlock = -1
        self.SetIcon(wx.Icon("Images/Polling.ico", wx.BITMAP_TYPE_ICO))
        self.listhosts = {}
        self.hiddenPasswd = "disabled"
        self.deencrypt_passwordword = "undefined"
        self.logtofile = logtofile
        self.logfile = logfile
        
        # Menu Fichier
        fileMenu = wx.Menu()
        fileMenu.Append(ID_IMPORTCSV, u"&Import CSV\tCTRL+i",
            u"Import CSV, semi colon separator ASCII file")
        fileMenu.Append(ID_EXPORTCSV, u"&Export CSV\tCTRL+e",
            u"Export CSV, semi colon separator ASCII file")
        fileMenu.AppendSeparator()
        fileMenu.Append(wx.ID_EXIT, u"&Exit\tCTRL+q", u"Exit PYNBM")
        
        #Menu Edit
        editMenu = wx.Menu()
        editMenu.Append(ID_SELECTALL, u"&Select All\tCTRL+a",u"Select All")        
        editMenu.Append(ID_COPY, u"&Copy\tCTRL+c",u"Copy Selection")
        editMenu.Append(ID_PASTE, u"&Paste\tCTRL+v",u"Paste")
        editMenu.Append(ID_SEARCH, u"&Find\tCTRL+f",u"Find")
        editMenu.Append(ID_SEARCHNEXT, u"&Find Next\tCTRL+g",u"Find Next")
        editMenu.Append(ID_SETTINGS, u"&Settings",u"Settings")        
        #editMenu.Append(ID_DEBUG, u"&debug\tCTRL+d",u"debug")                
        
        # Menu Afficher
        displayMenu = wx.Menu(style = wx.MENU_TEAROFF)
        toggleHidePasswordsMenu = displayMenu.AppendCheckItem(ID_TOGGLEHIDEPASSWD,
            u"&Hide passwords\tCTRL+h", u"Hide passwords")
        toggleHidePasswordsMenu.Check()
        self.toggleLockGrid = displayMenu.AppendCheckItem(ID_GRIDLOCK,
            u"&Lock the Grid\tCTRL+l", u"Lock the Grid in ReadOnly Mode")
        self.togglePolling = displayMenu.AppendCheckItem(ID_TOGGLEPOLLING,
            u"&Enable Polling\tCTRL+s", u"Enable Polling")
        self.resetstat = displayMenu.Append(ID_RESETSTAT,
            u"&Reset Statistics\tCTRL+r", u"Reset Statistics")
        self.resize = displayMenu.Append(ID_RESIZE,
            u"&Auto-resize cells", u"Auto-resize cells")
        # Menu A propos
        aboutMenu = wx.Menu(style = wx.MENU_TEAROFF)
        aboutMenu.Append(ID_ABOUT, u"&About PYNBM", u"About PYNBM")
        
        menuBar = wx.MenuBar()
        menuBar.Append(fileMenu, u"&File")
        menuBar.Append(editMenu, u"&Edit")
        menuBar.Append(displayMenu, u"&View")
        menuBar.Append(aboutMenu, u"&About")
        self.SetMenuBar(menuBar)
        
        # StatusBar
        self.bar = wx.StatusBar(self, -1)
        self.bar.SetFieldsCount(4)
        self.bar.SetStatusWidths([-1, -1, -1, -1])
        self.SetStatusBar(self.bar)
        
        # Toolbar
        self.tools = wx.ToolBar(self, -1, style = wx.TB_HORIZONTAL | wx.NO_BORDER)
        self.tools.AddSimpleTool(ID_IMPORTCSV,
            wx.Bitmap('Images/Import.png', wx.BITMAP_TYPE_PNG),
            shortHelpString = u"Import CSV",
            longHelpString = u"Import CSV, semi colon separator ASCII file")
        self.tools.AddSimpleTool(ID_EXPORTCSV, 
            wx.Bitmap('Images/Export.png', wx.BITMAP_TYPE_PNG),
            shortHelpString = u"Export CSV", 
            longHelpString = u"Export CSV, semi colon separator ASCII file")
        self.tools.AddSimpleTool(ID_SEARCH, 
            wx.Bitmap('Images/Search.png', wx.BITMAP_TYPE_PNG),
            shortHelpString = u"Search", 
            longHelpString = u"Search host or alias")

        self.tools.AddSimpleTool(ID_SETTINGS, 
            wx.Bitmap('Images/Settings.png', wx.BITMAP_TYPE_PNG),
            shortHelpString = u"Settings", longHelpString = u"Settings")
        self.pollingButton = self.tools.AddSimpleTool(ID_TOGGLEPOLLING, 
            wx.Bitmap('Images/Polling.png', wx.BITMAP_TYPE_PNG),
            shortHelpString = u"Toggle polling", 
            longHelpString = u"Toggle polling",
            isToggle=True)           
        self.gridlockButton = self.tools.AddSimpleTool(ID_GRIDLOCK,
            wx.Bitmap('Images/Locked.png', wx.BITMAP_TYPE_PNG),
            shortHelpString = u"Lock the grid", 
            longHelpString = u"Toggle grid edition mode",
            isToggle=True)          
        self.resetstatButton = self.tools.AddSimpleTool(ID_RESETSTAT,
            wx.Bitmap('Images/Resetstat.png', wx.BITMAP_TYPE_PNG),
            shortHelpString = u"Reset polling statistics", 
            longHelpString = u"Reset polling statistics")          
        self.resizeButton = self.tools.AddSimpleTool(ID_RESIZE,
            wx.Bitmap('Images/Resize.png', wx.BITMAP_TYPE_PNG),
            shortHelpString = u"Auto-resize cells", 
            longHelpString = u"Auto-resize cells")     

        
        self.tools.AddSeparator()
        self.tools.AddSimpleTool(wx.ID_EXIT,
            wx.Bitmap('Images/Exit.png', wx.BITMAP_TYPE_PNG),
            shortHelpString = u"Exit",
            longHelpString = u"Exit PYNBM")
        self.tools.Realize()
        self.SetToolBar(self.tools)

        # Handlers d'evenements
        wx.EVT_MENU(self, wx.ID_EXIT, self.on_exit)
        wx.EVT_MENU(self, ID_IMPORTCSV, self.on_import)
        wx.EVT_MENU(self, ID_COPY, self.on_copy)
        wx.EVT_MENU(self, ID_PASTE, self.on_paste)
        wx.EVT_MENU(self, ID_EXPORTCSV, self.on_export)
        wx.EVT_MENU(self, ID_TOGGLEPOLLING, self.on_toggle_polling)
        wx.EVT_MENU(self, ID_TOGGLEHIDEPASSWD, self.on_toggle_hide_passwords)
        wx.EVT_MENU(self, ID_SETTINGS, self.on_settings)
        wx.EVT_MENU(self, ID_GRIDLOCK, self.on_grid_lock)        
        wx.EVT_MENU(self, ID_RESETSTAT, self.on_resetstat)  
        wx.EVT_MENU(self, ID_RESIZE, self.on_resize)        
        wx.EVT_MENU(self, ID_ABOUT, self.on_about)
        wx.EVT_MENU(self, ID_SELECTALL, self.on_selectAll)
        wx.EVT_MENU(self, ID_SEARCH, self.on_search)
        wx.EVT_MENU(self, ID_SEARCHNEXT, self.on_search_next)        
        #wx.EVT_MENU(self, ID_DEBUG, self.on_debug)        
        self.Bind(wx.EVT_CLOSE, self.on_exit)
        self.Bind(EVT_UPDATE_RTT, self.rtt_update)
        self.Bind(EVT_UPDATE_CELL, self.updateCellValue)
        self.Bind(EVT_NEWVERSION, self.newVersionAvailable)
        self.Bind(EVT_THREAD_END, self.thread_status_update)
        self.Bind(EVT_EXTERNALTOOLSTATUS, self.external_tool_status)
        # sizer, instanciation de la grid
        sizer = wx.BoxSizer(wx.VERTICAL)
        self.pane = Grid(self)
        sizer.Add(self.pane, 1, wx.EXPAND|wx.ALL, 2)
        self.SetSizer(sizer)
        self.on_toggle_hide_passwords(self)

    def on_debug(self,evt):
        return

    def on_grid_lock(self,evt):
        """
        this function is called when the lock button is clicked.
        Toggles the grid edition mode
        """
        
        if (self.gridlock == 1): #unlock the grid if locked
            self.pane.unlock_edit()
            self.toggleLockGrid.Check(False)
            self.tools.ToggleTool(ID_GRIDLOCK,False)
            self.gridlock=0
            
        else:
            self.pane.lock_edit()
            self.toggleLockGrid.Check()
            self.tools.ToggleTool(ID_GRIDLOCK,True)
            self.gridlock=1
            
        
    def on_settings(self,evt):
        """
        this function is called when the settings button is clicked. creates
        a settings view by creating a Settings object
        """
        settingsview = Settings(self,-1, u"Settings", size=(350, 200))
        settingsview.CenterOnScreen()
        # this does not return until the dialog is closed.
        val = settingsview.ShowModal()
     
    def updateCellValue(self,evt):
        line = evt.itemline
        column = evt.itemcolumn
        value = evt.value
        self.pane.SetCellValue(int(line),int(column),value)
        

    def rtt_update(self,evt):
        """
        this function is called anytime a thread passes an RTT Update event
        to the MainWindow. Gets the item line and associtated round trip time
        and then update the grid with this values
        """
        line = evt.itemline
        rtt = evt.value
        host = evt.host
        
        self.pane.SetCellValue(line,6,rtt)
    
        statistics = self.pane.GetCellValue(line,7)
        
        if statistics != "" :
            success = int(statistics.split('/')[0])
            total = int(statistics.split('/')[1].split('-')[0])
        else:
            success = 0
            total = 0
        
        if re.search("^No Answer",rtt):
            if ((logtofile is True) and (logfile != None) and (logfile != "")):
                logger.info('%s did not answer', str(host))

            for i in range (0,len(self.pane.colnames)):
                if (i == 3):
                    continue
                self.pane.SetCellTextColour(line,i, wx.NamedColour('RED'))
                
        else:
            if ((logtofile is True) and (logfile != None) and (logfile != "")):
                logger.info('%s answered in %s', str(host), str(rtt))

            for i in range (0,len(self.pane.colnames)):  
                if (i == 3):
                    continue
                self.pane.SetCellTextColour(line,i, wx.NamedColour('BLACK'))
            success = success+1
        total = total+1
        percent = str(int(Decimal(success)/Decimal(total)*100))
        success = str(success)
        total = str(total)
        self.pane.SetCellValue(line,7, success+"/"+total+"-"+percent+"%")

    def on_toggle_hide_passwords(self,evt):
        """
        this function changes the color of password column to white so that
        they become invisible to the user. switches to black if prior status was
        hidden, and vice-versa.
        """
        if (self.hiddenPasswd == "disabled"):
            self.hide_passwd()
        else:
            self.show_passwd()
            
    def hide_passwd(self):
        self.hiddenPasswd = "enabled"
        lastlinenumber = self.get_last_line_number()
        for i in range(0,numberOfLines):
            self.pane.SetCellTextColour(i,3, wx.NamedColour('WHITE'))
        self.pane.ForceRefresh()            
    
    def show_passwd(self):
        self.hiddenPasswd = "disabled"
        lastlinenumber = self.get_last_line_number()
        for i in range(0,numberOfLines):
            self.pane.SetCellTextColour(i,3, wx.NamedColour('BLACK'))
        self.pane.ForceRefresh()

    def on_toggle_polling(self,evt,onexit=False):
        """
        when polling button is clicked, depending of the prior value of the
        Refresh class polling attribute, updated Refresh class polling 
        attribute and starts or stops polling threads
        """
        
        def set_on():
            #global threadlist
        
            pynbm_refresh.setPollingStatus("enabled")
            self.togglePolling.Check(True)
            self.tools.ToggleTool(ID_TOGGLEPOLLING,True)            
            # Instanciation du thread de refresh
            global interval
            global timeout

            for index in range(numberOfRefreshThreads):
                refreshthread = pynbm_refresh(self, 
                                        self.pane,UpdateRTTEvent,ThreadEndEvent,index,numberOfRefreshThreads,interval,timeout,ostype)
                refreshthread.start()
                time.sleep(0.01)
                self.bar.SetStatusText("Polling Threads count: " + 
                                       str(len(pynbm_refresh.threadlist)),2)
            
            
        def set_off():
            pynbm_refresh.setPollingStatus("locked")
            self.togglePolling.Check(False)
            self.tools.ToggleTool(ID_TOGGLEPOLLING,False)                        
            self.bar.SetStatusText(u"Polling Threads count : " +
                str(len(pynbm_refresh.threadlist)) + 
                u" finishing...",2)
                
                
                
        
        if (pynbm_refresh.getPollingStatus() == "disabled"):
            if (onexit == False):
                set_on()
        elif (pynbm_refresh.getPollingStatus() == "enabled"):
            set_off()
        else:
            pass        
        
        
    def external_tool_status(self,evt):
        self.bar.SetStatusText(evt.msg,3)    
            
    def thread_status_update(self,evt):
        """
        this function is called anytime a polling thread finishes, thus
        updating the status bar with the number of remaining polling threads
        """
        thread = evt.thread
        #global threadlist
        #print "finished thread " + str(thread)
        threadCount = len(pynbm_refresh.threadlist)
        
        finishing = ""
        if threadCount > 0:
            self.bar.SetStatusText(u"Polling Threads count : " +
                str(len(pynbm_refresh.threadlist)) + u" finishing... ",2)                
        else:
            pynbm_refresh.setPollingStatus("disabled")
            finishing = ""
            self.bar.SetStatusText(u"Polling Disabled",2)    
            
        
        

    def on_copy(self,evt):
        """
        CTRL+C
        """
        self.pane.copy()
    
    def on_paste(self,evt):
        """
        CTRL+V
        """
        self.pane.paste()
            
    def on_search(self,evt):
        """
        Search the grid for a pattern
        """
        self.pane.define_search()
    
    def on_selectAll(self,evt):
        self.pane.selectAll()

    def on_search_next(self,evt):
        """
        Search the grid for the next matching pattern
        """
        self.pane.search()

    def on_resetstat(self, evt):    
        """
        reset polling statistics (RTT and Rate)
        """
        for line in range(0,numberOfLines):
            for col in range(6,8):
                self.pane.SetCellValue(line,col,"")
            for col in range(len(self.pane.colnames)):
                if col == 3:
                    pass
                else:
                    self.pane.SetCellTextColour(line,col, wx.NamedColour('BLACK'))

    def on_resize(self, evt):    
        """
        resize grid cells to best fit to contents
        """
        self.pane.resize()
                
    def on_import(self, evt):
        """
        when import is clicked, prompt the use for a CSV file to import, and
        build the grid contents by reading the input file line by line
        """
        
        
        self.deencrypt_passwordword = "undefined"


        dlg = wx.FileDialog(self, 
                            u"Select a file", 
                            wildcard="All Files (*.*)|*.*", 
                            style = wx.OPEN) #wildcard = "*.*"
        dlgreturn = dlg.ShowModal()
        filepath = dlg.GetPath()
        filename = os.path.basename(filepath)
        dlg.Destroy()
        if dlgreturn == wx.ID_OK and filename != "":
            self.importfile(filepath)

    def importfile(self,filepath):
            
        for line in range(0,numberOfLines):
            for col in range(len(self.pane.colnames)):
                self.pane.SetCellValue(line,col,"")
        f = open(filepath, 'r')
        i = 0
        for line in f:
            j = 0
            cells = line.split(';')
            for col in cells:
                if (j==len(self.pane.colnames)):
                    break #change line if column number of CSV exceeds the grid's number of columns
                    #if password cell begins with "Encrypted:", then
                    #prompt for passphrase and decrypt the cell using it
                if ((j == 3) and (re.search("^Encrypted:.+",col))):
                    col = self.decrypt(col)
                self.pane.SetCellValue(i, j, col.rstrip('\n'))
                self.pane.SetCellTextColour(i,j, wx.NamedColour('BLACK'))
                j += 1
            i += 1
        if (self.hiddenPasswd == "enabled"):
            self.hide_passwd()
        filename = os.path.basename(filepath)
        self.bar.SetStatusText("Loaded File : " + filename,1)

        global lastOpenedFile
        lastOpenedFile = filepath
        myIniFile = IniFile(INI_FILENAME)
        myIniFile.save_to_file()

        self.SetTitle(self.titre + " - " + filepath)

    def decrypt(self,encrypted):
        if self.deencrypt_passwordword == "undefined":
            # Pop-up window prompting for decrypt password                        
            dlg = wx.PasswordEntryDialog(self, 
                 u"Please enter passphrase to decrypt password field :",
                 u"Enter passphrase")                          
            if dlg.ShowModal() == wx.ID_OK:  
                #  self.log.WriteText('You entered: %s\n' % dlg.GetValue())
                self.deencrypt_passwordword = dlg.GetValue()
                dlg.Destroy()
        encrypted = encrypted.replace("Encrypted:","")
        cryptObj = pynbm_aes.AESOperations()
        decrypted = cryptObj.aes_decrypt(encrypted,self.deencrypt_passwordword)
        return decrypted

    def get_last_line_number(self):
        """
        this function parses up the grid from the last line and stop when
        finding a non-empty cell in column 6 or 0 and returns the current
        line number
        """
        #seek for line number of last non-empty line:
        for linenumber in range(numberOfLines-1,-1,-1):
            if (self.pane.GetCellValue(linenumber,0) != "" or 
               self.pane.GetCellValue(linenumber,6) != ""):
                linenumber = linenumber + 1
                break
        return linenumber
        


    def on_export(self, evt):
        """
        when export is selected, prompt the user for a file to which the grid
        will be dumped
        """
        dlg = wx.FileDialog(self, 
                            u"Select a file", 
                            wildcard = "All Files (*.*)|*.*", 
                            style = wx.SAVE)
        dlgreturn = dlg.ShowModal()
        filepath = dlg.GetPath()
        filename = dlg.GetFilename()
        dlg.Destroy()
        if dlgreturn == wx.ID_OK and filename != "":
            if os.path.exists(filepath):
                dlg = wx.MessageDialog(self, 
                                      u"File already exists, Are you sure you want to overwrite ?",
                                      u"File already exists",
                                      wx.YES_NO | wx.ICON_INFORMATION
            #wx.YES_NO | wx.NO_DEFAULT | wx.CANCEL | wx.ICON_INFORMATION
                                      )
                dlgreturn = dlg.ShowModal()
                dlg.Destroy()
                if dlgreturn == wx.ID_NO:
                    return
            
            enencrypt_passwordword1 = ""
            enencrypt_passwordword2 = ""
           
            # Pop-up window prompting for encrypt password                        
            dlg1 = wx.PasswordEntryDialog(self, 
                      u"Please enter passphrase to encrypt password " +
                      u" field (leave empty for no encryption):",
                      u"Enter passphrase")
#            dlg.SetValue("")
            if dlg1.ShowModal() == wx.ID_OK:
                enencrypt_passwordword1 = dlg1.GetValue()
            dlg1.Destroy()
            
            # Pop-up window prompting once again for encrypt password                        
            dlg2 = wx.PasswordEntryDialog(self, 
                      u"Please enter passphrase once again " +
                      u"(leave empty fo no encryption):",
                      u"Enter passphrase")
#            dlg.SetValue("")
            if dlg2.ShowModal() == wx.ID_OK:
                enencrypt_passwordword2 = dlg2.GetValue()
            dlg2.Destroy()
        
            if (enencrypt_passwordword1 != enencrypt_passwordword2):
                dlg = wx.MessageDialog(self, 
                                       u"passpharases do not match, aborting",
                                       u"passphrase check failed",
                                       wx.OK | wx.ICON_INFORMATION
                #wx.YES_NO | wx.NO_DEFAULT | wx.CANCEL | wx.ICON_INFORMATION
                                       )
                dlg.ShowModal()
                dlg.Destroy()
        
            else:      
                f = open(filepath, 'w')
                i = 0
                j = 0
                lastlinenumber = self.get_last_line_number()
                for i in range(0,lastlinenumber):
                    line = []
                    for j in range(len(self.pane.colnames)):
                        cell = self.pane.GetCellValue(i, j)
                        if ((j == 3) and (enencrypt_passwordword2 != "")):
                            cryptObj = pynbm_aes.AESOperations()
                            cell = u"Encrypted:" + cryptObj.aes_encrypt(cell,enencrypt_passwordword2)
                        line.append(cell)
                    i += 1
                    j = 0
                    line = ";".join(line)
                    f.write(line + '\n')
                f.close()
                self.SetTitle(self.titre + " - " + filepath)

    def on_exit(self, evt):
        """
        when exit is clicked, set Refresh class polling attribute to disabled,
        thus informing all threads to exit their polling loop and close.
        Wait for all polling threads to close, and exit the application
        """

        self.on_toggle_polling(None,True)
        
        dlg = wx.MessageDialog(self, 
                              u"Do you want to save the grid before exiting ?",
                              u"Save the grid before exiting ?" ,
                              wx.YES_NO | wx.ICON_INFORMATION
    #wx.YES_NO | wx.NO_DEFAULT | wx.CANCEL | wx.ICON_INFORMATION
                              )
        dlgreturn = dlg.ShowModal()
        dlg.Destroy()
        if dlgreturn == wx.ID_YES:
            save = self.on_export(self)
        
        #global threadlist
        
        while len(pynbm_refresh.threadlist) > 0:
            time.sleep(0.1)
            self.bar.SetStatusText(u"Polling Threads count : " +
                str(len(pynbm_refresh.threadlist)) + u" finishing... ",2)  
        time.sleep(0.1)
        self.Destroy()

    def on_about(self,evt):
        """
        function called when about is clicked from the menu, show the about
        dialog informations
        """
        # First we create and fill the info object
        info = wx.AboutDialogInfo()
        info.Name = Name
        info.Version = Version + " - " + Release_Date
        info.Copyright = u"(C) 2012 Florian Pervès"
        info.Description = wordwrap(
            u"PYNBM - Python Network Bookmark Manager - is a simple tool "
            u"that focuses on sytems or network devices remote access. "
            u"PYNBM gives the ability to launch any network client tool "
            u"upon systems defined in the grid, and this is only one "
            u"right-click away. Enjoy ! :)",
            350, 
            wx.ClientDC(self))
        info.WebSite = (u"http://florian.perves.free.fr/PYNBM", 
                        u"PYNBM HomePage")
        info.Developers = [ u"Florian Pervès" ]
        licenseText = u"Open Source, freely redistributable"
        info.License = wordwrap(licenseText, 500, wx.ClientDC(self))

        # Then we call wx.AboutBox giving it that info object
        wx.AboutBox(info)

    def newVersionAvailable(self,evt):
    
        self.currentversion = evt.currentVersion
        dlg = wx.MessageDialog(self, 
                              u"PyNBM " + self.currentversion + " is available, do you want to download it right now ?",
                              u"PyNBM " + self.currentversion + " is available !" ,
                              wx.YES_NO | wx.ICON_INFORMATION
    #wx.YES_NO | wx.NO_DEFAULT | wx.CANCEL | wx.ICON_INFORMATION
                              )
        dlgreturn = dlg.ShowModal()
        dlg.Destroy()
        if dlgreturn == wx.ID_YES:
            webbrowser.open("%s" % downloadURL, new = 0)
            return

            
                    
class Settings(wx.Dialog):
    def __init__( self, parent, ID, title, size=wx.DefaultSize, pos=wx.DefaultPosition,
            style=wx.DEFAULT_DIALOG_STYLE, useMetal=False):
        """
        function that builds the settings view
        """

        provider = wx.SimpleHelpProvider()
        wx.HelpProvider.Set(provider)
        
        # Instead of calling wx.Dialog.__init__ we precreate the dialog
        # so we can set an extra style that must be set before
        # creation, and then we create the GUI object using the Create
        # method.


        pre = wx.PreDialog()
        pre.SetExtraStyle(wx.DIALOG_EX_CONTEXTHELP)

        pre.Create(parent, ID, title, pos, size, style)

        # This next step is the most important, it turns this Python
        # object into the real wrapper of the dialog (instead of pre)
        # as far as the wxPython extension is concerned.
        self.PostCreate(pre)

        # Now continue with the normal construction of the dialog
        # contents
        
        sizer = wx.BoxSizer(wx.VERTICAL)
        box = wx.BoxSizer(wx.HORIZONTAL)
        global interval
        label = wx.StaticText(self, -1, u"Interval (seconds):")
        label.SetHelpText(u"This defines the interval in seconds between " +
                          u"two pollings")
        box.Add(label, 0, wx.ALIGN_CENTRE|wx.ALL, 5)
        text = wx.TextCtrl(self, -1, str(interval), size=(80,-1))
        text.SetHelpText(u"Enter an integer or float value, " +
                         u"use dot as decimal separator")
        self.input1=text
        box.Add(text, 1, wx.ALIGN_CENTRE|wx.ALL, 5)
        sizer.Add(box, 0, wx.GROW|wx.ALIGN_CENTER_VERTICAL|wx.ALL, 5)
        
        global timeout
        box = wx.BoxSizer(wx.HORIZONTAL)
        label = wx.StaticText(self, -1, u"Timeout (seconds):")
        label.SetHelpText(u"This defines the ping timeout in seconds")
        box.Add(label, 0, wx.ALIGN_CENTRE|wx.ALL, 5)
        text = wx.TextCtrl(self, -1, str(timeout), size=(80,-1))
        text.SetHelpText(u"Enter an integer value")
        self.input2=text
        box.Add(text, 1, wx.ALIGN_CENTRE|wx.ALL, 5)
        sizer.Add(box, 0, wx.GROW|wx.ALIGN_CENTER_VERTICAL|wx.ALL, 5)

        box = wx.BoxSizer(wx.HORIZONTAL)
        label = wx.StaticText(self, -1, u"Polling threads :")
        label.SetHelpText(u"This defines the number of parallel " +
                          u"polling threads to be launched")
        box.Add(label, 0, wx.ALIGN_CENTRE|wx.ALL, 5)
        text = wx.TextCtrl(self, -1,str(numberOfRefreshThreads), size=(80,-1))
        text.SetHelpText(u"Enter an integer value. Lauching too many " +
                         u"threads (>20) can lead to high CPU/Memory " +
                         u"consumption, and even to a freeze of the " +
                         u"application. A value of 10 might be convenient " +
                         u"in most of the cases.")
        self.input3 = text
        box.Add(text, 1, wx.ALIGN_CENTRE|wx.ALL, 5)     
        sizer.Add(box, 0, wx.GROW|wx.ALIGN_CENTER_VERTICAL|wx.ALL, 5)

        line = wx.StaticLine(self, -1, size=(20,-1), style = wx.LI_HORIZONTAL)
        sizer.Add(line, 0, wx.GROW|wx.ALIGN_CENTER_VERTICAL|wx.RIGHT|wx.TOP, 5)
        
        box = wx.BoxSizer(wx.HORIZONTAL)
        cb1 = wx.CheckBox(self, -1, u"Enable logging")
        self.enablelogging = cb1
        cb1.SetHelpText(u"Enable file logging of ping threads")
        if (logtofile is True):
            cb1.SetValue(True)
        box.Add(cb1, 0, wx.ALIGN_CENTRE|wx.ALL, 5)
        sizer.Add(box, 0, wx.GROW|wx.ALIGN_CENTER_VERTICAL|wx.ALL, 5)        

        box = wx.BoxSizer(wx.HORIZONTAL)        
        text = wx.TextCtrl(self, -1, str(logfile), size=(80,-1))
        text.SetHelpText(u"Path to Log File")
        self.input4=text
        box.Add(text, 1, wx.ALIGN_CENTRE|wx.ALL, 5)
        sizer.Add(box, 0, wx.GROW|wx.ALIGN_CENTER_VERTICAL|wx.ALL, 5)
        
        box = wx.BoxSizer(wx.HORIZONTAL)
        browsebtn = wx.Button(self, wx.ID_FILE, "Browse")
        browsebtn.SetHelpText(u"Browse for file to which write logs")
        browsebtn.SetDefault()
        box.Add(browsebtn,0, wx.ALIGN_CENTRE|wx.ALL, 5)
        sizer.Add(box, 0, wx.GROW|wx.ALIGN_CENTER_VERTICAL|wx.ALL, 5)        
        
        line = wx.StaticLine(self, -1, size=(20,-1), style = wx.LI_HORIZONTAL)
        sizer.Add(line, 0, wx.GROW|wx.ALIGN_CENTER_VERTICAL|wx.RIGHT|wx.TOP, 5)
        
        box = wx.BoxSizer(wx.HORIZONTAL)
        label = wx.StaticText(self, -1, u"Grid Lines count (needs restart) :")
        label.SetHelpText(u"This defines the number of lines " +
                          u"in the grid")
        box.Add(label, 0, wx.ALIGN_CENTRE|wx.ALL, 5)
        text = wx.TextCtrl(self, -1,str(numberOfLines), size=(80,-1))
        text.SetHelpText(u"Enter an integer value.")
        self.input5 = text
        box.Add(text, 1, wx.ALIGN_CENTRE|wx.ALL, 5)     
        sizer.Add(box, 0, wx.GROW|wx.ALIGN_CENTER_VERTICAL|wx.ALL, 5)
        
        box = wx.BoxSizer(wx.HORIZONTAL)
        label = wx.StaticText(self, -1, u"Custom Columns, semicolon (;) separator (needs restart) :")
        label.SetHelpText(u"This defines custom columns " +
                          u"in the grid.")
        box.Add(label, 0, wx.ALIGN_CENTRE|wx.ALL, 5)
        sizer.Add(box, 0, wx.GROW|wx.ALIGN_CENTER_VERTICAL|wx.ALL, 5)

        box = wx.BoxSizer(wx.HORIZONTAL)
        text = wx.TextCtrl(self, -1,str(";".join(customColumns)), size=(80,-1))
        text.SetHelpText(u"Enter an string.")
        self.input6 = text
        sizer.Add(box, 0, wx.GROW|wx.ALIGN_CENTER_VERTICAL|wx.ALL, 5)
        box.Add(text, 1, wx.ALIGN_CENTRE|wx.ALL, 5)     

        box = wx.BoxSizer(wx.HORIZONTAL)
        
        line = wx.StaticLine(self, -1, size=(20,-1), style = wx.LI_HORIZONTAL)
        sizer.Add(line, 0, wx.GROW|wx.ALIGN_CENTER_VERTICAL|wx.RIGHT|wx.TOP, 5)
        
        extensionbtn = wx.Button(self, 1, "Customize Extensions")
        self.Bind(wx.EVT_BUTTON, self.on_extension, extensionbtn)
        extensionbtn.SetDefault()
        extensionbtn.SetHelpText(u"Edit extensions cofiguration file")
        box.Add(extensionbtn,0, wx.ALIGN_CENTRE|wx.ALL, 5)
        sizer.Add(box, 0, wx.GROW|wx.ALIGN_CENTER_VERTICAL|wx.ALL, 5)        
        
        line = wx.StaticLine(self, -1, size=(20,-1), style = wx.LI_HORIZONTAL)
        sizer.Add(line, 0, wx.GROW|wx.ALIGN_CENTER_VERTICAL|wx.RIGHT|wx.TOP, 5)

        box = wx.BoxSizer(wx.HORIZONTAL)
        cb3 = wx.CheckBox(self, -1, u"Load last opened file on startup")
        self.loadLastOpenedFileOnStartup = cb3
        cb3.SetHelpText(u"Load last opened file on startup")
        if (loadLastOpenedFileOnStartup is True):
            cb3.SetValue(True)
        box.Add(cb3, 0, wx.ALIGN_CENTRE|wx.ALL, 5)
        sizer.Add(box, 0, wx.GROW|wx.ALIGN_CENTER_VERTICAL|wx.ALL, 5)      
        
        box = wx.BoxSizer(wx.HORIZONTAL)
        cb2 = wx.CheckBox(self, -1, u"Check for updates on startup")
        self.checkForUpdate = cb2
        cb2.SetHelpText(u"Check for updates on startup")
        if (checkForUpdate is True):
            cb2.SetValue(True)
        box.Add(cb2, 0, wx.ALIGN_CENTRE|wx.ALL, 5)
        sizer.Add(box, 0, wx.GROW|wx.ALIGN_CENTER_VERTICAL|wx.ALL, 5)        

        
        line = wx.StaticLine(self, -1, size=(20,-1), style = wx.LI_HORIZONTAL)
        sizer.Add(line, 0, wx.GROW|wx.ALIGN_CENTER_VERTICAL|wx.RIGHT|wx.TOP, 5)
        
        
        btnsizer = wx.StdDialogButtonSizer()

        if wx.Platform != "__WXMSW__":
            btn = wx.ContextHelpButton(self)
            btnsizer.AddButton(btn)


        okbtn = wx.Button(self, wx.ID_OK)
        okbtn.SetHelpText(u"The OK button completes the dialog")
        okbtn.SetDefault()
        btnsizer.AddButton(okbtn)

        cancelbtn = wx.Button(self, wx.ID_CANCEL)
        cancelbtn.SetHelpText(u"The Cancel button cancels the dialog. " +
                              u"(Cool, huh?)")
        btnsizer.AddButton(cancelbtn)
        btnsizer.Realize()


        sizer.Add(btnsizer, 0, wx.ALIGN_CENTER_VERTICAL|wx.ALL, 5)

        self.SetSizer(sizer)
        sizer.Fit(self)
        self.Bind(wx.EVT_CHECKBOX, self.evt_check_box, cb1)
        self.Bind(wx.EVT_BUTTON, self.on_button_browse, browsebtn)
        self.Bind(wx.EVT_BUTTON, self.on_button_ok, okbtn)
        self.Bind(wx.EVT_BUTTON, self.on_button_cancel, cancelbtn)







        
    def on_button_browse(self, evt):
        dlg = wx.FileDialog(self, 
                            u"Select a file", 
                            wildcard="All Files (*.*)|*.*", 
                            style = wx.SAVE)
        dlgreturn = dlg.ShowModal()
        self.filepath = dlg.GetPath()
        self.filename = dlg.GetFilename()
        self.input4.SetValue(self.filepath)
        dlg.Destroy()
    
    def on_extension(self,evt):
        """
        edit the extensions file
        """
        try:
            f = open(EXTENSIONS_FILENAME, 'r')
        except IOError:
            shutil.copy2('Extensions.defaults', EXTENSIONS_FILENAME)
            self.on_extension(evt)
        else:
            if (ostype == "win32"):
                command = "start "+ "\"\" " + "\"" + EXTENSIONS_FILENAME + "\""
                #os.system("start "+EXTENSIONS_FILENAME)
            elif (ostype == "darwin"):
                command = "open " + "\"" + EXTENSIONS_FILENAME + "\""
            elif (ostype == "linux2"):
                command = "xdg-open " + "\"" + EXTENSIONS_FILENAME + "\""
            
            subprocess.call(command, shell=True)
        

    def evt_check_box(self, event):
        #print 'evt_check_box: %d\n' % event.IsChecked()        
        pass

    def on_button_ok(self, evt):
        """
        when OK button is clicked in settings view, update Refresh class
        attributes, and close
        """
        
        def RepresentsInt(s):
            try: 
                int(s)
                return True
            except ValueError:
                return False
        
        def RepresentsFloat(s):
            try: 
                float(s)
                return True
            except ValueError:
                return False
        
        if RepresentsInt(self.input1.GetValue()):
            global interval
            interval = int(abs(float(self.input1.GetValue())))
            pynbm_refresh.setInterval(interval)
        else:
            dlg = wx.MessageDialog(self, 
                                  u"Interval must be an integer number",
                                  u"Invalid interval value",
                                  wx.OK | wx.ICON_INFORMATION
            #wx.YES_NO | wx.NO_DEFAULT | wx.CANCEL | wx.ICON_INFORMATION
                                  )
            dlg.ShowModal()
            dlg.Destroy()
            return
        
        if RepresentsInt(self.input2.GetValue()):
            global timeout
            timeout = int(abs(int(self.input2.GetValue())))
            pynbm_refresh.setTimeout(timeout)
        else:
            dlg = wx.MessageDialog(self, 
                                   u"Timeout must be an integer number",
                                   u"Invalid timeout value",
                                   wx.OK | wx.ICON_INFORMATION
            #wx.YES_NO | wx.NO_DEFAULT | wx.CANCEL | wx.ICON_INFORMATION
                                   )
            dlg.ShowModal()
            dlg.Destroy()
            return

        if RepresentsInt(self.input3.GetValue()):
            global numberOfRefreshThreads
            numberOfRefreshThreads = int(abs(int(self.input3.GetValue())))
        else:
            dlg = wx.MessageDialog(self, 
                                   u"Polling threads must be an integer number",
                                   u"Invalid polling threads value",
                                   wx.OK | wx.ICON_INFORMATION
            #wx.YES_NO | wx.NO_DEFAULT | wx.CANCEL | wx.ICON_INFORMATION
                                   )
            dlg.ShowModal()
            dlg.Destroy()
            return

        if ((self.enablelogging.GetValue() is True) and 
           ((self.input4.GetValue() == "None") or 
            (self.input4.GetValue() == ""))):
            dlg = wx.MessageDialog(self, 
                                   u"If you want to log to a file, you need " +
                                   u"to specify in which file...!",
                                   u"No logfile selected",
                                   wx.OK | wx.ICON_INFORMATION
            #wx.YES_NO | wx.NO_DEFAULT | wx.CANCEL | wx.ICON_INFORMATION
                                   )
            dlg.ShowModal()
            dlg.Destroy()
            return

        if RepresentsInt(self.input5.GetValue()):
            global numberOfLines
            numberOfLines = int(abs(int(self.input5.GetValue())))
        else:
            dlg = wx.MessageDialog(self, 
                                   u"Number of Lines must be an integer number",
                                   u"Invalid number of line value",
                                   wx.OK | wx.ICON_INFORMATION
            #wx.YES_NO | wx.NO_DEFAULT | wx.CANCEL | wx.ICON_INFORMATION
                                   )
            dlg.ShowModal()
            dlg.Destroy()
            return

        global customColumns
        customColumns = self.input6.GetValue().split(";")
        
        global logtofile
        logtofile = self.enablelogging.GetValue()

        
        global logfile
        logfile = self.input4.GetValue()

        global loadLastOpenedFileOnStartup
        loadLastOpenedFileOnStartup = self.loadLastOpenedFileOnStartup.GetValue()

        global checkForUpdate
        checkForUpdate = self.checkForUpdate.GetValue()
        
        myIniFile = IniFile(INI_FILENAME)
        myIniFile.save_to_file()
    
        self.Destroy()

    def on_button_cancel(self, evt):
        """
        when Cancel button is clicked in settings view, close window
        """
        self.Destroy()
        
class MyApp(wx.App):
    def __init__(self):
        myIniFile = IniFile(INI_FILENAME)
        myIniFile.load_from_file()
        wx.App.__init__(self, False)
        mainWindowTitle = "%s - %s" % (Name,Version)
        window = MainWindow(mainWindowTitle)
        window.Show(True)
        self.SetTopWindow(window)
        
        #update checker run in a separate thread, and thus needs to first
        # be instantiated, and then started with the start() method
        if checkForUpdate == True:
            checkUpdate = pynbm_updateChecker(Version,window,NewVersion,checkVersionURL)
            checkUpdate.start()
        if loadLastOpenedFileOnStartup == True:
            if os.path.exists(lastOpenedFile):
                window.importfile(lastOpenedFile)
            
        return None

app = MyApp()
app.MainLoop()
