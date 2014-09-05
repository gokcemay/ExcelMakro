Attribute VB_Name = "Module2"
Sub FrictionGraph()



Dim Incrmnt, LoadNo As String
Dim MyFile, Title, MyMacroFile, MyTextFile, Numune As String
Dim MyRow, MyRow2, Veri, MyValue2, MyValue3, MyColumn, MyColumn2, iRow, iColumn, LastRow As Long
Dim RngDis As Range
Dim RngFri As Range
Dim sh As Worksheet
Dim chrt As Chart

Set sh = ActiveWorkbook.Worksheets("Sayfa3")
Set chrt = sh.Shapes.AddChart.Chart


Sheets("Data").Select ' Data Sheetini seç
iColumn = Worksheets("Data").Cells(2, Columns.Count).End(xlToLeft).Column + 1 'son dolu columndan bir sonrakini seç



Application.ScreenUpdating = False
' Name current file
MyMacroFile = ActiveWorkbook.Name
' Prompt for file
MyFile = Application.GetOpenFilename("All Files,*.*")
If MyFile = False Then
Exit Sub
End If

LoadNo = "Distance [m]"
' Open file
Workbooks.OpenText Filename:=MyFile, Origin:=xlWindows, StartRow:=1, _
    DataType:=xlDelimited, Tab:=True, FieldInfo:=Array(0, 2)
' Name text file
MyTextFile = ActiveWorkbook.Name
' Find cell with "Distance"
Do
    Windows(MyTextFile).Activate
        ' Exit loop if can't find any matches
    On Error GoTo Err_Fix
        Cells.Find(What:=LoadNo, After:=ActiveCell, LookIn _
            :=xlValues, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:= _
            xlNext, MatchCase:=False).Activate
    On Error GoTo 0
    ' Exit loop if starting search over form top
    If ActiveCell.Row <= MyRow Then
        Exit Do
    End If

MyRow = ActiveCell.Row
MyRow2 = ActiveCell.Row + 1


MyColumn = ActiveCell.Column
MyColumn2 = ActiveCell.Column + 3


Set sh = ActiveSheet


'Veri = sh.Cells(MyRow, MyColumn)   Artýk kullanýlmýyor
'Veri2 = sh.Cells(MyRow2, MyColumn2)


Set RngDis = Range(Cells(MyRow, MyColumn), Cells(MyRow, MyColumn).End(xlDown)) 'Distance dahil sonuna kadar seç Ayrýca Object olarak tanýmlýyorum böylece propertylerini kullanabiliyorum value gibi :)
Set RngFri = Range(Cells(MyRow, MyColumn2), Cells(MyRow, MyColumn2).End(xlDown)) 'Laps dahil sonuna kadar seç
LastRow = Cells(MyRow, MyColumn).End(xlDown).Row        '65536 Satýr kaydetmemek için son satýrý bulma
MsgBox LastRow

   Windows(MyMacroFile).Activate
    Sheets("Data").Select

iRow = Worksheets("Data").Cells(Rows.Count, iColumn).End(xlUp).Row + 1 'son dolu satýrý seç

Numune = MyFile             'Numune deðiþkenine dosyanýn ismini ata ki üstüne yazabileyim

    'Numune ismini almak için bir loop
    Do
    Numune = Mid(Numune, InStr(1, Numune, "\") + 1, InStr(1, Numune, ".") - InStr(1, Numune, "\")) '\ ile . arasýný al
    i = InStr(1, Numune, "\")
    Loop Until i = 0

Numune = Mid(Numune, 1, InStr(1, Numune, ".") - 1) 'Numune isminde . kaldý onu da temizle


Range(Cells(iRow - 1, iColumn), Cells(iRow - 1, iColumn)) = Numune  'Ýlk boþ satýr ve kolona numune yaz
Worksheets("Data").Range(Cells(iRow, iColumn), Cells(LastRow, iColumn)).Value = RngDis.Value    'ilk boþ sütuna distance gir
Worksheets("Data").Range(Cells(iRow, iColumn + 1), Cells(LastRow, iColumn + 1)) = RngFri.Value  'ikinci sütuna laps gir




Workbooks.OpenText Filename:=MyFile, Origin:=xlWindows, StartRow:=1, _
    DataType:=xlDelimited, Tab:=True, FieldInfo:=Array(0, 2)




Loop

Err_Fix:
Windows(MyTextFile).Activate
ActiveWorkbook.Close
Application.ScreenUpdating = True

Exit Sub



End Sub
