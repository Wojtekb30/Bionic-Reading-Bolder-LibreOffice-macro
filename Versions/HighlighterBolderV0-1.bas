REM  *****  BASIC  *****

Sub Highlighter
doc = thisComponent
selected = doc.CurrentSelection.getByIndex(0).getString()
arrayy=Split(selected)
text=""
For Each item in arrayy
dlugosc = len(item)
dlugoscs = int(len(item)/2)
wone = MID(item, 1, dlugoscs)
wtwo = MID(item, dlugoscs+1, dlugosc)
text = text + Ucase(wone) + wtwo + " "
Next
doc.CurrentSelection.getByIndex(0).SetString(text)
MsgBox("Done")
End Sub

