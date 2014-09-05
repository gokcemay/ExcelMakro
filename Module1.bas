Attribute VB_Name = "Module1"
Sub CF_AL()
Attribute CF_AL.VB_Description = "Dosyadan CF bilgilerini alýp iþleyecek"
Attribute CF_AL.VB_ProcData.VB_Invoke_Func = " \n14"
'
' CF_AL Makro
' Dosyadan CF bilgilerini alýp iþleyecek
'

'

Dim Incrmnt, LoadNo As String
Dim MyFile, Title, MyMacroFile, MyTextFile, Numune As String
Dim MyRow, MyValue2, MyValue3 As Long
Dim i As Integer
i = 1

Application.ScreenUpdating = False
' Name current file
MyMacroFile = ActiveWorkbook.Name
' Prompt for file
MyFile = Application.GetOpenFilename("All Files,*.*")
If MyFile = False Then
Exit Sub
End If

LoadNo = "Start"
' Open file
Workbooks.OpenText Filename:=MyFile, Origin:=xlWindows, StartRow:=1, _
    DataType:=xlDelimited, Tab:=True, FieldInfo:=Array(0, 2)
' Name text file
MyTextFile = ActiveWorkbook.Name
' Find cell with "Start"
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
' Get last 12 characters of 12 row below
Title = ActiveCell.Offset(0, 3)
'MsgBox Title 'Kontrol için koymuþtum
MyValue2 = ActiveCell.Offset(1, 3)
'MsgBox MyValue2 'Kontrol için koymuþtum
MyRow = ActiveCell.Row
' Paste value in spreadsheet with macro in columns A and B
Windows(MyMacroFile).Activate
    Range("A65536").End(xlUp).Offset(1, 0) = MyFile
    
    Range("B65536").End(xlUp).Offset(1, 0) = Title
Range("C65536").End(xlUp).Offset(1, 0) = MyValue2
Numune = MyFile

    'Numune ismini almak için bir loop
    Do
    Numune = Mid(Numune, InStr(1, Numune, "\") + 1, InStr(1, Numune, ".") - InStr(1, Numune, "\")) '\ ile . arasýný al
    i = InStr(1, Numune, "\")
    Loop Until i = 0

Numune = Mid(Numune, 1, InStr(1, Numune, ".") - 1) 'Numune isminde . kaldý onu da temizle

Range("D65536").End(xlUp).Offset(1, 0) = Numune

Loop

Err_Fix:
Windows(MyTextFile).Activate
ActiveWorkbook.Close
Application.ScreenUpdating = True

Exit Sub
End Sub

