from distutils.core import setup
import py2exe

setup(windows=[
                {
                    "script" : "PYNBM.py",
                    "icon_resources": [(0, "Images\polling.ico")]
                }
               ],
       options={
                "py2exe":{
                        "dll_excludes": ["mswsock.dll", "POWRPROF.dll"]
                    }
                }
)

