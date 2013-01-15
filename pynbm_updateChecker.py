#!/usr/bin/python -tt
# -*- coding: utf-8 -*-
import threading
import urllib2
import wx.lib.newevent

class pynbm_updateChecker(threading.Thread):
    def __init__(self,Version,window,newVersionEvent,checkVersionURL):
        threading.Thread.__init__(self)
        self.Version = Version
        self.window = window
        self.NewVersion = newVersionEvent
        self.checkUpdateURL = checkVersionURL
        
    def getCurrentVersion(self):
        proxy_handler = urllib2.ProxyHandler()
        opener = urllib2.build_opener(proxy_handler)
        urllib2.install_opener(opener)
        try:
            f = urllib2.urlopen(self.checkUpdateURL)
            version = f.read(10)
        except:
            try:
                proxy_handler = urllib2.ProxyHandler({})
                opener = urllib2.build_opener(proxy_handler)
                urllib2.install_opener(opener)    
                f = urllib2.urlopen(self.checkUpdateURL)
                version = f.read(10)
            except:
                version = 'unknown'
        return version

    def run(self):
        currentVersion = self.getCurrentVersion()
        if (currentVersion != 'unknown' and currentVersion != self.Version):
                evt = self.NewVersion(currentVersion = currentVersion)
                wx.PostEvent(self.window, evt)