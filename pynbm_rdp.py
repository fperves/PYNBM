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

import win32crypt
import subprocess


class RDP():
    """
    Class that creates an RDP file and loads it with mstsc.exe
    """
    def __init__(self,host,domain,login,password,rdpfilename,timer=5):
        try:
            f = open(rdpfilename,'w')
        except IOError:
            return
        else:
            contents = "screen mode id:i:2\n\
desktopwidth:i:1280\n\
desktopheight:i:800\n\
session bpp:i:32\n\
winposstr:s:0,1,0,0,1024,510\n\
compression:i:1\n\
keyboardhook:i:2\n\
displayconnectionbar:i:1\n\
disable wallpaper:i:1\n\
disable full window drag:i:1\n\
allow desktop composition:i:0\n\
allow font smoothing:i:0\n\
disable menu anims:i:1\n\
disable themes:i:0\n\
disable cursor setting:i:0\n\
bitmapcachepersistenable:i:1\n\
full address:s:"+ host + "\n\
audiomode:i:0\n\
redirectprinters:i:1\n\
redirectcomports:i:0\n\
redirectsmartcards:i:1\n\
redirectclipboard:i:1\n\
redirectposdevices:i:0\n\
drivestoredirect:s:*\n\
autoreconnection enabled:i:1\n\
authentication level:i:0\n\
prompt for credentials:i:0\n\
negotiate security layer:i:1\n\
username:s:" + login + "\n"
            if (domain != ""):
                    contents = contents + "domain:s:" + domain + "\n"
            contents = contents + "password 51:b:" + self.encrypt_password(password) + "\n\
remoteapplicationmode:i:0\n\
alternate shell:s:\n\
shell working directory:s:\n\
gatewayhostname:s:\n\
gatewayusagemethod:i:4\n\
gatewaycredentialssource:i:4\n\
gatewayprofileusagemethod:i:0\n\
promptcredentialonce:i:1"
            f.write(contents)    
            f.close()
            commandline = "mstsc.exe %s" % rdpfilename
            args = commandline.split(' ')
            pid = subprocess.Popen(args).pid
            time.sleep(timer)
            os.unlink(rdpfilename)

    def encrypt_password(self,password):
        pwdHashStr = ""
        unicodestring = password.encode("utf-16-le")
        pwdHash = win32crypt.CryptProtectData(unicodestring,'psw',None,None,None,0)
        for char in pwdHash:
        # conversion of each char in hexadecimal on 2 digits with 0 padding if 
        # necessary ("02" : 00, 01, ... 0d, 0e, 0f)
            pwdHashStr += "%02X" % ord(char) 
        return pwdHashStr            
