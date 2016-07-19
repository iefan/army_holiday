#!/usr/local/bin/python
# -*- coding: utf-8 -*-
# Used successfully in Python2.5 with matplotlib 0.99 and wxpython 2.8.9.1
from distutils.core import setup
import py2exe
import sys


# no arguments

if len(sys.argv) == 1:
    sys.argv.append("py2exe")
# We need to import the glob module to search for all files.
#import glob
import matplotlib as mpl

# We need to exclude matplotlib backends not being used by this executable.  You may find
# that you need different excludes to create a working executable with your chosen backend.
# We also need to include include various numerix libraries that the other functions call.

#  "includes" : ["matplotlib.backends",  "matplotlib.backends.backend_qt4agg",
#               "matplotlib.figure","pylab", "numpy", "matplotlib.numerix.fft",]
#               "matplotlib.numerix.linear_algebra", "matplotlib.numerix.random_array",
#               "matplotlib.backends.backend_tkagg"]
opts = {
    'py2exe': { "includes" : ["matplotlib.backends.backend_tkagg"],
#                'excludes': ['_gtkagg', '_tkagg', '_agg2', '_cairo', '_cocoaagg',
#                             '_fltkagg', '_gtk', '_gtkcairo', ],
#                'dll_excludes': ['libgdk-win32-2.0-0.dll',
#                                 'libgobject-2.0-0.dll'],
                "compressed": 1,"ascii": 0,
              }
       }


# note: using matplotlib.get_mpldata_info
data_files = mpl.get_py2exe_datafiles()

# for console program use 'console = [{"script" : "scriptname.py"}]
setup(windows=[{"script" : "frmlogin.pyw", \
            "icon_resources": [(0, "bitmap/PHRLogo.ico")]}], \
    options=opts,  \
    zipfile = None, \
    data_files=data_files,\
    name="PHRecord")
