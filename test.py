#from win32com.client import Dispatch
#
#myWord = Dispatch('Word.Application')
#myWord.Visible = 1  # comment out for production
#
#myWord.Documents.Open(working_file)  # open file
#
## ... doing something
#
#myWord.ActiveDocument.SaveAs(new_file)
#myWord.Quit() # close Word.Application


import win32com.client as win32
word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = True
doc = word.Documents.Add()
doc.PageSetup.TextColumns.SetCount (2)

rng = doc.Range(0,0)
rng.Text = "Lots and lots of words.  " * 200

# Collapse the range so we point at the end.

rng.Collapse( win32.constants.wdCollapseEnd)

# Insert a hard page break.

# https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2003/aa211923(v%3Doffice.11)
rng.InsertBreak( win32.constants.wdPageBreak )

# Insert more words.

rng.Text = "More words.  " * 30

doc.SaveAs('test.doc')
