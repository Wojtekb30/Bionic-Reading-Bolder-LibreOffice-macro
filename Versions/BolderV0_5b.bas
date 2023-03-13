REM  *****  BASIC  *****

Sub BolderV0_5b

REM Bionic Reading soluton created by Woj, 06.03.2023
REM V0.5b

doc = ThisComponent
selected = doc.CurrentSelection


dlugoscdoc = int(InputBox ("Click where to begin and enter numer of letters to process (processing may take long, don't edit or interact with document until ''Done'' message appears):","Bolder start"))
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


