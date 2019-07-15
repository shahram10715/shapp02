# -*- coding: utf-8 -*-

def wcontrols(path):
    etabs974 = open(path, 'a')
    maintxt = ''
    maintxt += '$ CONTROLS\n'
    maintxt += '  UNITS  "TON"  "M"  \n'
    maintxt += '  PREFERENCE  MERGETOL 0.001\n'
    maintxt += '  RLLF  METHOD "INFLAREAASCE795"  USEDEFAULTMIN "YES"  \n' 
    maintxt += '\n'
    etabs974.write(maintxt)
    etabs974.close()

