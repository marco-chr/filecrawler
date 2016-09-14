# encoding=utf8
__author__ = 'marco'
# This script reads word 97-03 .doc file properties, revision, filename, dirname and compiles info into an excel file
# Maximum depth is 2 which means ./level1/level2/filename.doc

import os, string
import pypyodbc
import xlsxwriter
import zipfile
import time
import re
import pythoncom
from win32com import storagecon
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

# These come from ObjIdl.h
FMTID_UserDefinedProperties = "{F29F85E0-4FF9-1068-AB91-08002B27B3D9}"

PIDSI_TITLE               = 0x00000002
PIDSI_SUBJECT             = 0x00000003
PIDSI_AUTHOR              = 0x00000004
PIDSI_COMMENTS            = 0x00000006
PIDSI_CREATE_DTM          = 0x0000000c

def PrintStats(filename) :
    title=''
    subject=''
    comments=''
    if sys.platform == 'win32':
        if not pythoncom.StgIsStorageFile(filename) :
            print "The file is not a storage file!"
            return title, subject, comments
        # Open the file.
        flags = storagecon.STGM_READ | storagecon.STGM_SHARE_EXCLUSIVE
        stg_= pythoncom.StgOpenStorage(filename, None, flags )

        # Now see if the storage object supports Property Information.
        try:
            pss = stg_.QueryInterface(pythoncom.IID_IPropertySetStorage)
        except pythoncom.com_error:
            print "No summary information is available"
            return title, subject, comments
        # Open the user defined properties.
        ps = pss.Open(FMTID_UserDefinedProperties)
        props = PIDSI_TITLE, PIDSI_SUBJECT, PIDSI_AUTHOR, PIDSI_COMMENTS, PIDSI_CREATE_DTM
        data = ps.ReadMultiple( props )
        # Unpack the result into the items.
        title, subject, author, comments, created = data        
        return title, subject, comments

workbook = xlsxwriter.Workbook('NOVARTIS_HU.xlsx')
worksheet = workbook.add_worksheet()
row = 0
path = '.'
path = os.path.normpath(path)
res = []

for dirname, dirnames, filenames in os.walk(path, topdown=True):
    # print path to all subdirectories first.
    depth = dirname[len(path) + len(os.path.sep):].count(os.path.sep)

    if depth == 1:
        for filename in filenames:
            if filename.endswith('.doc'):
                command = 'C:\\antiword\\antiword ' + '"' + dirname + "\\" + filename + '"' + ' > file.txt'
                os.system(command)
                time.sleep(1)
                ufunction=''
                revs=[]
                toggle_start=False
                toggle_stop=False

                for line in open('file.txt','r').readlines():
                    m = re.search('Fonction:\s+(.+)\n',line)
                    if m:
                        ufunction = m.group(1)
                        break

                for line2 in open('file.txt','r').readlines():
                    if toggle_start==False:
                        n = re.search('\|s         \|',line2)
                        if n:
                            toggle_start=True
                    elif toggle_start==True:
                        n = re.search('\|(\d+)\s+\|',line2)
                        if n:
                            revs.append(n.group(1))
                        elif re.search('\|          \|',line2):
                            break

                curfile = dirname + '\\' + filename
                print curfile
                root,area,unit = dirname.split('\\')
                worksheet.write(row, 0, area.decode('utf-8','ignore'))
                worksheet.write(row, 1, unit.decode('utf-8','ignore'))
                worksheet.write(row, 2, filename.decode('latin-1','ignore'))
                [title, subject, comments]=PrintStats(curfile)
                if subject:
                    worksheet.write(row, 3, subject.decode('latin-1','ignore'))
                if comments:
                    worksheet.write(row, 4, comments.decode('latin-1','ignore'))
                if len(revs) > 0:
                    worksheet.write(row, 5, max(revs))
                row += 1

workbook.close()
