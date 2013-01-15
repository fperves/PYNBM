#!/usr/bin/python -tt
# -*- coding: utf-8 -*-
import re
# WxPython libraries
import wx.lib.newevent
import threading
import time
import pynbm_wmiping
import os

class pynbm_refresh(threading.Thread):
    polling = "disabled"
    interval = ""
    timeout = ""
    threadlist = []
    
    def __init__(self,panel,gridview,UpdateRTTEvent,ThreadEndEvent,threadIndex,nbOfThreads,interval,timeout,ostype):
        """
        Refresh class maps to an individual thread that is in charge of pinging
        hosts and updating round trip times in grid depending on the configuration,
        several refresh objects (and so threads) can be created to distribute
        the pinging tasks and update the grid rtt fields faster
        """
        threading.Thread.__init__(self)
        if (ostype == "darwin") : self.osping=getattr(self, "darwin_ping")
        elif (ostype == "win32") : self.osping=getattr(self, "windows_ping")
        elif (ostype == "linux2") : self.osping=getattr(self, "linux_ping")
        self.panel = panel
        self.gridview = gridview
        self.threadIndex = threadIndex
        self.nbOfThreads = nbOfThreads
        self.ThreadEndEvent = ThreadEndEvent
        self.UpdateRTTEvent = UpdateRTTEvent
        self.lock = threading.Lock()
        pynbm_refresh.setInterval(interval)
        pynbm_refresh.setTimeout(timeout)

    @staticmethod
    def getPollingStatus():
        return pynbm_refresh.polling
        
    @staticmethod
    def setPollingStatus(status):
        pynbm_refresh.polling = status
        
    @staticmethod
    def getInterval():
        return pynbm_refresh.interval

    @staticmethod
    def setInterval(interval):
        pynbm_refresh.interval = interval

    @staticmethod
    def getTimeout():
        return pynbm_refresh.timeout
        
    @staticmethod
    def setTimeout(timeout):
        pynbm_refresh.timeout = timeout
        
    def darwin_ping(self,host):
         """
         specific MacOS function for pinging hosts
         """
         p = os.popen("ping -q -c 2 -t "+ str(pynbm_refresh.getInterval()) +
                      " -W " + str(pynbm_refresh.getInterval()) + " " + host + 
                      " 2>&1")
         answer = p.read()
         p.close()
         res = re.search("( 100.0% packet loss)",answer)
         if res:
             return ('No Answer')
         else:
             if len(answer.split('/')) < 4:
                 return ('No Answer')
             else:
                 rtt = answer.split('/')[4]
                 unit = answer.split(' ')[-1].rstrip('\n')
                 return (rtt+" "+unit)

    def linux_ping(self,host):
         """
         specific Linux function for pinging hosts
         """
         p = os.popen("ping -q -c 2 -W " + str(pynbm_refresh.getInterval()) + " " + host + " 2>&1")
         answer = p.read()
         p.close()
         res = re.search("( 100.0% packet loss)",answer)
         if res:
             return ('No Answer')
         else:
             if len(answer.split('/')) < 4:
                 return ('No Answer')
             else:
                 rtt = answer.split('/')[4]
                 unit = answer.split(' ')[-1].rstrip('\n')
                 return (rtt+" "+unit)

    def windows_ping(self,host):
        """
        specific Windows function for pinging hosts
        """        
        StatusCodeDef = {0: 'Success',
            11001: 'Buffer Too Small',
            11002: 'Destination Net Unreachable',
            11003: 'Destination Host Unreachable',
            11004: 'Destination Protocol Unreachable',
            11005: 'Destination Port Unreachable',
            11006: 'No Resources',
            11007: 'Bad Option',
            11008: 'Hardware Error',
            11009: 'Packet Too Big',
            11010: 'Request Timed Out',
            11011: 'Bad Request',
            11012: 'Bad Route',
            11013: 'TimeToLive Expired Transit',
            11014: 'TimeToLive Expired Reassembly',
            11015: 'Parameter Problem',
            11016: 'Source Quench',
            11017: 'Option Too Big',
            11018: 'Bad Destination',
            11032: 'Negotiating IPSEC',
            11050: 'General Failure',
            None: 'Unknown Error'}
        
        p = pynbm_wmiping.WMIPing(host, str(pynbm_refresh.getTimeout()))
        res = p.Ping()
        code = res[0]
        value = res[1]
        if code == 0:
            #success
            return str(value) + " ms"
        elif code == 1:
            #name resolution failure
            return ('No Answer : Name Resolution Failure')
        elif code == 2:
            #other causes
            return ('No Answer : ' + StatusCodeDef[value])
        
        
    def run(self):
        """
        the run method is called by the start() method inherited from Threading
        module, which effectively starts the thread.
        Each thread loops on pinging the hosts it has to query sequentially, 
        and send RTTEvent to the MainWindow.
        They exit loop when Refresh class polling attribute switches to 
        disabled.
        """
        #global threadlist
        
        self.lock.acquire()
        self.__class__.threadlist.append(self)
        #print len(threadlist)
        #print "started thread " + str(self)
        self.lock.release()
        while (True):
            if (pynbm_refresh.getPollingStatus() != "enabled"):
                break
            lastlinenumber = self.panel.get_last_line_number()
            line = 0

            # ping loop testing each host
            for line in range(self.threadIndex,lastlinenumber, 
                              self.nbOfThreads): 
                if (pynbm_refresh.getPollingStatus() != "enabled"):
                    break
                #ex with 10 threads : th1 deals with items on lines 0,10,20,30... 
                #th2 deals with items on lines 1,11,21,31, etc
                if self.gridview.GetCellValue(line,0) == "":
                    continue
                # test if item is disabled
                if self.gridview.GetCellValue(line,5) == "disabled": 
                    continue
                host = self.gridview.GetCellValue(line,0)
                if (host.count(':') == 1):
                    host = host.split(":")[0]
                #print str(self) + " pinging line " + str(line) + ":" + str(host)
                rtt = str(self.osping(host))
                evt = self.UpdateRTTEvent(itemline = line, host = host, value = rtt)
                wx.PostEvent(self.panel, evt)

            # sleeping loop
            for i in range(0,int(pynbm_refresh.getInterval())): #sleep N * 1 seconds
                if (pynbm_refresh.getPollingStatus() != "enabled"):
                    break
                time.sleep(1)
                
        evt = self.ThreadEndEvent(thread = self)
        wx.PostEvent(self.panel, evt)
        self.lock.acquire()
        self.__class__.threadlist.remove(self)
        self.lock.release()