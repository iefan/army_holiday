from distutils.core import setup
import py2exe
import sys


# no arguments

if len(sys.argv) == 1:
    sys.argv.append("py2exe")

# creates a standalone .exe file, no zip files
setup( options = {"py2exe": {"compressed": 1, "optimize": 2, "ascii": 0, \
    "bundle_files": 1}},
    zipfile = None,
    # replace myFile.py with your own code filename here ...
    windows = [{"script": 'frmlogin.pyw',
                "icon_resources": [(0, "bitmap/PHRLogo.ico")]     ### Icon to embed into the PE file.
                }] )
