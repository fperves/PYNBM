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
#!/usr/bin/python -tt
# -*- coding: utf-8 -*-

import win32com.client
import pythoncom

class WMIPing():
    def __init__(self,host,timeout):
        self.host = host
        self.timeout = int(timeout)*1000
        self.timeout = str(self.timeout)

    
    def Ping(self):        
        pythoncom.CoInitialize()
        self.wmiobj = win32com.client.GetObject(r"winmgmts:\\.\root\cimv2")
        col_items = self.wmiobj.ExecQuery("Select * from Win32_PingStatus Where Address='%s' and timeout=%s" % (self.host , self.timeout))
        
        for item in col_items:
            if item.StatusCode == 0:
                # success
                res = [0, str(item.ResponseTime)]
            elif item.StatusCode == None:
                # failure in name resolution
                res = [1, item.PrimaryAddressResolutionStatus]
            else:    
                # other failure code
                res = [2, item.StatusCode]

        pythoncom.CoUninitialize()
        return res
