#!c:\Python27\python.exe
#MysqlHa Relationships   
#MysqlHeartbeat Relationships   
#ReplikorReplication Relationships
#ReadWriteDb Relationships

import Tkinter,tkFileDialog
import ctypes  # An included library with Python install.
def Mbox(title, text, style):
	rc = ctypes.windll.user32.MessageBoxW(0, text, title, style)
	return rc
##  Styles:
##  0 : OK
##  1 : OK | Cancel
##  2 : Abort | Retry | Ignore
##  3 : Yes | No | Cancel
##  4 : Yes | No
##  5 : Retry | No 
##  6 : Cancel | Try Again | Continue

## IDABORT 3
## IDCANCEL 2
## IDCONTINUE 11
## IDIGNORE 5
## IDNO 7
## IDOK 1
## IDRETRY 4
## IDTRYAGAIN 10
## IDYES 6


root = Tkinter.Tk()
root.withdraw()
#filename = tkFileDialog.askopenfile(parent=root,mode='r',filetypes=[("PDF file","*.pdf")],title='Choose an SO PDF file')
filename = tkFileDialog.askopenfilename(filetypes=[("PDF file","*.pdf")],title='Choose an SO PDF file')
if filename != None:
	responseno = Mbox(u'File selected', filename, 1)
	if ( responseno == 2 ):
		exit()
print responseno
print "OK"
    #print "This PDF file has been selected", filename
