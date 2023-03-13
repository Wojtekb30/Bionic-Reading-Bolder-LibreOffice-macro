REM  *****  BASIC  *****

Sub BolderV0_6_fullDoc

REM Bionic Reading soluton created by Woj, 06.03.2023
REM V0.6

doc = ThisComponent
selected = doc.CurrentSelection
MsgBox("Click in document from where in text to format, then click OK to start macro. Processing may take long so don't close, change or interact with document until the macro is done working and ''Done'' message appears.")
oVC = thisComponent.getCurrentController.getViewCursor

cursor = oVC.getText.createTextCursorByRange(oVC)	
oVC.gotoStart(false)
oVC.gotoEnd(true)
dlugoscdoc = len(oVC.string)
oVC.gotoStart(false)
if dlugoscdoc<1 then
dlugoscdoc=1
end if

wordlen = 0
n = 0
Do Until n = dlugoscdoc
cursor.goRight(1, true)

if n=0 or n=1 then
cursor.CharWeight = 200
wordlen = wordlen+1
end if

if n>1 then
cursor.CharWeight = 100
if cursor.isEndOfWord() then
cursor.goLeft(int(wordlen/2), true)
cursor.CharWeight = 200
cursor.goRight(int(wordlen/2), false)
wordlen = -1
end if
wordlen = wordlen + 1
end if

n=n+1
Loop

MsgBox("Done. Thank you for using my macro!")
End Sub


