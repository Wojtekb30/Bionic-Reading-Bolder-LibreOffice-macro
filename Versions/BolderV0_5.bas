REM  *****  BASIC  *****

Sub BolderV0_5

REM Bionic Reading soluton created by Woj, 06.03.2023
REM V0.5

doc = ThisComponent
selected = doc.CurrentSelection


dlugoscdoc = int(InputBox ("Click where to begin and enter numer of letters to process:","Bolder start"))
if dlugoscdoc<1 then
dlugoscdoc=1
end if

oVC = thisComponent.getCurrentController.getViewCursor

cursor = oVC.getText.createTextCursorByRange(oVC)	
oVC.gotoStart(false)

wordlen = 0
n = 0
Do Until n = dlugoscdoc
cursor.goRight(1, true)
currenttext = cursor.String

if n=0 or n=1 then
cursor.CharWeight = 200
wordlen = wordlen+1
end if

if n>1 then
dlugosczaznaczenia = len(currenttext)
lastchar = MID(currenttext, dlugosczaznaczenia, dlugosczaznaczenia)
cursor.CharWeight = 100
if lastchar = " " or lastchar = "" or cursor.isStartOfWord() then
cursor.goLeft(int(wordlen/2)+1, true)
cursor.CharWeight = 200
cursor.goRight(int(wordlen/2)+1, false)
wordlen = 0-1
end if
wordlen = wordlen + 1
end if

n=n+1
Loop

MsgBox("Done")
End Sub


