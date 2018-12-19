import win32com.client
import win32com

ppt = win32com.client.Dispatch('PowerPoint.Application')
ppt.Presentations.Open(r'C:\Users\name\Desktop\Test.pptm')

# Exported Module Directory
ppt.VBE.ActiveVBProject.VBComponents.Import(r'C:\Users\name\Desktop\Module1.bas')

# Name of Sub
ppt.Run('Test')

ppt.Quit()
del ppt
