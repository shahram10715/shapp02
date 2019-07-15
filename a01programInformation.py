# -*- coding: utf-8 -*-

def wprogramInformation(path):
    etabs974 = open(path, 'a')
    maintxt = ''
    maintxt += '$ PROGRAM INFORMATION\n'
    maintxt += 'PROGRAM  "ETABS"  VERSION "9.7.4"  \n'
    maintxt += '\n'
    etabs974.write(maintxt)
    etabs974.close()
