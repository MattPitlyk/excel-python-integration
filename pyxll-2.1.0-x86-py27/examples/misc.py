"""
PyXLL Examples: Misc

Utility functions used in the examples spreadsheet but not
part of any particular example

www.pyxll.com
"""
import sys
import os
import pyxll
from pyxll import xl_func

@xl_func(": string", volatile=True)
def pyxll_version():
    """returns the pyxll version"""
    return pyxll.__version__

@xl_func(": string", volatile=True)
def python_version():
    """returns the python version"""
    return ".".join([str(x) for x in sys.version_info])

@xl_func(": bool", volatile=True)
def win32api_is_installed():
    """returns True if win32api can be imported"""
    try:
        import win32api
        return True
    except ImportError:
        return False
        
@xl_func(": bool", volatile=True)
def win32com_is_installed():
    """returns True if win32com is installed"""
    try:
        import win32com
        return True
    except ImportError:
        return False

@xl_func(": bool", volatile=True)
def numpy_is_installed():
    """returns True if numpy is installed"""
    try:
        import numpy
        return True
    except ImportError:
        return False    

@xl_func(": string", volatile=True)
def pyxll_logfile():
    """returns the pyxll logfile path"""
    config = pyxll.get_config()
    if config.has_option("LOG", "path") and config.has_option("LOG", "file"):
        path = os.path.join(config.get("LOG", "path"), config.get("LOG", "file"))
        return path
    return "No log path set in the PyXLL config"
    
@xl_func("xl_cell: string", macro=True)
def get_formula(cell):
    """returns the formula of a cell"""
    return cell.formula

@xl_func("xl_cell: string", macro=True)
def get_array_formula(cell):
    """returns the formula of a cell"""
    return "{%s}" % cell.formula