#!/usr/bin/python -tt
# -*- coding: utf-8 -*-
import re
# WxPython libraries
import wx
import wx.lib.newevent
import subprocess
import threading

class pynbm_CustomContextItem(threading.Thread):
    def __init__(self,mainwindow,UpdateColEvent,ExternalToolStatusEvent):
        """
        create objets items in right-click from extensions.csv file, and maps
        custom events binding to launch the appropriate extension
        """
        threading.Thread.__init__(self)
        self.mainwindow=mainwindow #mainwindow to post returned value
        self.UpdateColEvent=UpdateColEvent #event to call to update cell with returned value
        self.ExternalToolStatusEvent=ExternalToolStatusEvent
       

        
    def create_item(self,funcname,commandline,menu,grid,subMenuName=None):

        self.funcname = funcname
        self.commandline = commandline        
        self.menu = menu
        self.grid = grid
        self.subMenuName = subMenuName

        res = None
        if (self.funcname == "Context_Menu_Separator"):
            self.menu.AppendSeparator()
        elif (self.funcname == "subMenuEnd"):
            self.menu.AppendMenu(wx.NewId(), str(self.subMenuName), self.commandline)
        elif (self.funcname == "subMenu"):
            res = wx.Menu()
        else:
            funcname = self.funcname
            if re.search("^\$\{.+\}:(.+)",self.funcname):
                funcname = re.search("^\$\{.+\}:(.+)",self.funcname).group(1)
            var = wx.NewId()
            self.contextitem = self.menu.Append(var,funcname)
            self.grid.Bind(wx.EVT_MENU, self.custom_function, self.contextitem)
            
            res = self.contextitem
        return res
        
    def custom_function(self,evt):
        self.start()
       
    def run(self):
        """
        custom_function is called when a menu item is clicked in grid's contextual 
        menu
        """
        if (self.commandline != ""):
            #count selected non-empty items
            count=0
            for row in (self.grid.get_selected_rows()):
                if str(self.grid.GetCellValue(row,0)) == "":
                    #skip this row if host is empty
                    continue
                else:
                    count=count+1
                    
            print "non-empty rows count : " + str(count)
            statuscount=0
            
            try:
                functionlabel=self.funcname.split(":")[1]
            except:
                functionlabel=self.funcname
           
            status = functionlabel + " starting..."
            print status
            evt1=self.ExternalToolStatusEvent(msg = status)
            wx.PostEvent(self.mainwindow, evt1)
         
            for row in (self.grid.get_selected_rows()):
                
                if str(self.grid.GetCellValue(row,0)) == "":
                    #skip this row if host is empty
                    continue
                statuscount=statuscount+1
                commandline = self.commandline

                for i in range(len(self.grid.colnames)):
                    i=i+1
                    pattern = "${%i}" % i            
                    commandline = commandline.replace(pattern,
                        str(self.grid.GetCellValue(row,i-1)))

                args = commandline.split(' ')

                if re.search("^\$\{(.+)\}:",self.funcname): #if line starts with ${xx} : care for the returned value

                    
                    try:
                        startupinfo = subprocess.STARTUPINFO()
                        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                        returnedString = subprocess.check_output(args,startupinfo=startupinfo,stderr=subprocess.STDOUT).rstrip()
                    except subprocess.CalledProcessError as error:
                        returnedString = error.output

                               
                    
                    destColId=int((re.search("^\$\{(.+)\}:",self.funcname)).group(1))-1
                    evt = self.UpdateColEvent(itemline = row, itemcolumn = destColId, value = returnedString)
                    wx.PostEvent(self.mainwindow, evt)
                else:
                    #print "args " + str(args)
                    pid = subprocess.Popen(args).pid
                    
                if (statuscount == count):
                    status = functionlabel + " finished : " + str(statuscount) + "/" + str(count)
                else:
                    status = functionlabel + " running : " + str(statuscount) + "/" + str(count)
                evt2=self.ExternalToolStatusEvent(msg = status)
                wx.PostEvent(self.mainwindow, evt2)
                #print status
                    