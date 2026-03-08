# writer/uno_utils.py
import uno
from com.sun.star.beans import PropertyValue


def inches(n): return int(n * 2540)
def pt(n):     return int(n * 35.3)

def prop(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p

def fixed_ls(h):
    ls = uno.createUnoStruct("com.sun.star.style.LineSpacing")
    ls.Mode = 1
    ls.Height = h
    return ls

def prop_ls(pct):
    ls = uno.createUnoStruct("com.sun.star.style.LineSpacing")
    ls.Mode = 0
    ls.Height = pct
    return ls