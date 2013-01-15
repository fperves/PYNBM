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

###############################################################################
# AES symetric key encryption and decryption functions
from base64 import b64encode, b64decode
from Crypto.Cipher import AES
import hashlib
import re

class AESOperations():
    def aes_encrypt(self,plaintext,passphrase):
        m = hashlib.md5()						
        m.update(passphrase)
        passphrase = m.hexdigest()                #hash md5 (32 chars)
        aes = AES.new(passphrase, AES.MODE_CFB)
        #text to encrypt
        plaintext = "*****$$$$$PYNBM Crypt$$$$$*****%s" % plaintext
        #b64 transforms encrypted string to asci chars.
        ciphertext = b64encode(aes.encrypt(plaintext))       
        return ciphertext

    def aes_decrypt(self,ciphertext,passphrase):
        m = hashlib.md5()
        m.update(passphrase)
        passphrase = m.hexdigest()
        aes = AES.new(passphrase, AES.MODE_CFB)
        decrypted = aes.decrypt(b64decode(ciphertext))
        res = re.search ("^\*\*\*\*\*\$\$\$\$\$PYNBM Crypt\$\$\$\$\$\*\*\*\*\*",
                        decrypted)
        if res:
            return decrypted.replace("*****$$$$$PYNBM Crypt$$$$$*****","")
        else:
            return 'wrong passphrase'