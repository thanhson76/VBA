'Alt + F11 để vào VBE --> menu Tools --> chọn References ... --> tìm và chọn "Microsoft Scripting Runtime" --> nhấn OK

Dim FSO As New FileSystemObject
Dim F As Folder
Dim SF As Folder
Dim File As File
Dim ListPath(1 To 100000, 1 To 6) As String
Dim k As Long
Dim lastRow As Long
Sub Batdau()
'
' Tao 3 nut bam
'

'
    ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, 400, 100, 100, _
        25).Select
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "Ch" & ChrW(7885) & "n th" & ChrW(432) & " m" & ChrW(7909) & "c"
    Selection.OnAction = "Getfile"
    ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, 400, 130, 100, _
        25).Select
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = ChrW(272) & ChrW(7893) & "i tên file"
        Selection.OnAction = "Rename"
    ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, 400, 160, 100, _
        25).Select
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "Xoá"
       Selection.OnAction = "Clear"
    Range("A1").Select
End Sub
Sub Getfile()
    'Tao ten A1:G1
    Columns("A:A").ColumnWidth = 5
    Columns("B:B").ColumnWidth = 28
    Columns("C:C").ColumnWidth = 30
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "STT"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "" & ChrW(272) & ChrW(432) & ChrW(7901) & "ng d" & ChrW(7851) & "n"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Tên File/Folder"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Ðuôi file"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "File/Folder"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "L" & ChrW(7847) & "n truy c" & ChrW(7853) & "p cu" & ChrW(7889) & "i"
    Columns("D:F").EntireColumn.AutoFit
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "Tên m" & ChrW(7899) & "i"
    Range("A1:G1").Select
    Selection.Font.Bold = True
    Columns("D:G").EntireColumn.AutoFit
    
    'To dam va mau
    Range("A1:G1").Select
    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
        ' Lay du lieu
    Range("A2:G100000").Clear
    Dim FD As FileDialog
    Set FD = Application.FileDialog(msoFileDialogFolderPicker)
    k = 1
    If FD.Show Then
        Set F = FSO.GetFolder(FD.SelectedItems(1))
        DoFolder F
        
'        For Each SF In F.SubFolders
'            'xuat thong tin folder
''            ListPath(k, 1) = k  'STT
''            ListPath(k, 2) = SF.ParentFolder.Path
''            ListPath(k, 3) = SF.Name  'Name
''            ListPath(k, 4) = ""  'Extension
''            ListPath(k, 5) = "Folder"  'FileType
''            ListPath(k, 6) = SF.DateLastAccessed  'DateLastAccessed
''            k = k + 1
'            DoFolder SF
'        Next
        
        'Xuat ket qua
        Range("A2").Resize(k, 6) = ListPath
    End If
    

    
    'Xoa dong thua
    ' Tìm ô cu?i cùng ch?a d? li?u trong c?t A
    
    'Dim lastRow As Long CHO CAI NAY LEN TREN
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Xóa d?nh d?ng c?a các ô sau ô cu?i cùng ch?a d? li?u
    
    ws.Range("A" & lastRow + 1 & ":A" & ws.Rows.Count).ClearFormats

    ' Luu b?ng tính
    ' ThisWorkbook.Save
    ' Cap nhat lai bang tinh
    ActiveSheet.UsedRange
    'Co dinh hang 1 va gian dong cot F
    Range("A1").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.AutoFilter
    Rows("2:2").Select
    ActiveWindow.FreezePanes = True
    Columns("F").EntireColumn.AutoFit
    
End Sub

Sub DoFolder(Folder As Folder)
    Dim SubFolder As Folder
    For Each SubFolder In Folder.SubFolders
        DoFolder SubFolder
    Next
    
    For Each File In Folder.Files
        ListPath(k, 1) = k  'STT
        ListPath(k, 2) = File.ParentFolder.Path  'Path
        ListPath(k, 3) = FSO.GetBaseName(File.Path)  'Name
        ListPath(k, 4) = "." & FSO.GetExtensionName(File.Path)  'FileType
        ListPath(k, 5) = "File"  'FileType
        ListPath(k, 6) = File.DateLastAccessed  'DateLastAccessed
        k = k + 1
    Next
End Sub
Sub Rename()
    Dim A As Variant
    Dim Path1 As String, Path2 As String
    A = Range("A1").CurrentRegion.Value
    If IsArray(A) Then
        Cells.Interior.Color = xlNone
        On Error Resume Next
        For k = 2 To UBound(A, 1)
            If A(k, 7) <> "" Then
                If A(k, 5) = "Folder" Then
                    'Doi ten Folder se anh huong den file
'                    Path1 = A(k, 2) & "\" & A(k, 3)   'duong dan cu
'                    Path2 = A(k, 2) & "\" & A(k, 7)   'duong dan moi
'                    'Name Path1 As Path2
'                    FSO.MoveFolder Path1, Path2
                    
                Else
                    Path1 = A(k, 2) & "\" & A(k, 3) & A(k, 4)   'duong dan cu
                    Path2 = A(k, 2) & "\" & A(k, 7) & A(k, 4)   'duong dan moi
                    'Name Path1 As Path2
                    FSO.CopyFile Path1, Path2, True
                    FSO.DeleteFile Path1
                End If
            End If
            'Neu doi ten khong duoc thi bao do
            If Err.Number > 0 Then
                Err.Clear
                Range("A" & k).Resize(1, 7).Interior.Color = 255
            End If
        Next
        On Error GoTo 0
    End If
    MsgBox "Xong", vbInformation, ""
End Sub

Sub Clear()
    Range("A2:G100000").Clear
End Sub
