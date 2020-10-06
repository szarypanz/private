Sub PCReportFormat()
    
    Dim fileString As String

    fileString = Application.GetOpenFilename()
    
    If fileString <> "False" Then
    
        Dim wb As Workbook
    
        Set wb = Workbooks.Open(FileName:=fileString)
        Set wb = ActiveWorkbook

        hideunusedcells 'calling procedure for hiding rows w/o data
    
        report_formatting 'calling procedure for copying formats by row
        
        MsgBox "done"
        ActiveWorkbook.Worksheets("Introduction").Activate
        Range("A1").Select
        ActiveWorkbook.Close SaveChanges:=True

    Else

        Exit Sub    'closing sub if no file was chosen
  
    End If

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub hideunusedcells() 'hiding rows w/o data
   
Dim ws As Worksheet

    For Each ws In ActiveWorkbook.Worksheets 'goes by each worksheet, all tabs were prepared with this in mind
            
        lrow = ws.Cells(Rows.Count, "A").End(xlUp).Row
        Range(ws.Cells(lrow + 2, 1), ws.Cells(Rows.Count, Columns.Count)).EntireRow.Hidden = True
    
    Next ws
    
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub report_formatting() 'copying formats by row - each tab is done separately in case dimensions change etc

Dim rangecopy As Range, rangepaste As Range


ActiveWorkbook.Worksheets("Summary").Activate 'copying formats for Summary tab
With ActiveSheet
    Set rangecopy = .Range(.Range("A2"), .Cells(2, Columns.Count).End(xlToLeft))
    Set rangepaste = .Range(.Range("A2"), .Cells(Rows.Count, 1).End(xlUp)).Resize(, rangecopy.Columns.Count)
End With
rangecopy.Copy
rangepaste.PasteSpecial Paste:=xlPasteFormats
Application.CutCopyMode = False
Range("A1").Select


' rest of the tabs
ActiveWorkbook.Worksheets("Excluded Entities").Activate 'copying formats for Excluded Entities tab
If Range("A5") > 0 Then                                 ' (If) statement allows script to skip empty tabs
    With ActiveSheet
        Set rangecopy = .Range(.Range("A5"), .Cells(5, Columns.Count).End(xlToLeft))
        Set rangepaste = .Range(.Range("A5"), .Cells(Rows.Count, 1).End(xlUp)).Resize(, rangecopy.Columns.Count)
    End With
    rangecopy.Copy
    rangepaste.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    Range("A1").Select
End If


ActiveWorkbook.Worksheets("Excluded Incumbents").Activate 'copying formats for Excluded Incumbents tab
If Range("A5") > 0 Then
    With ActiveSheet
        Set rangecopy = .Range(.Range("A5"), .Cells(5, Columns.Count).End(xlToLeft))
        Set rangepaste = .Range(.Range("A5"), .Cells(Rows.Count, 1).End(xlUp)).Resize(, rangecopy.Columns.Count)
    End With
    rangecopy.Copy
    rangepaste.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    Range("A1").Select
End If


ActiveWorkbook.Worksheets("AGC Rematches").Activate 'copying formats for AGC Rematches tab
If Range("A5") > 0 Then
    With ActiveSheet
        Set rangecopy = .Range(.Range("A5"), .Cells(5, Columns.Count).End(xlToLeft))
        Set rangepaste = .Range(.Range("A5"), .Cells(Rows.Count, 1).End(xlUp)).Resize(, rangecopy.Columns.Count)
    End With
    rangecopy.Copy
    rangepaste.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    Range("A1").Select
End If


ActiveWorkbook.Worksheets("Multiple YT Positions").Activate 'copying formats for Multiple YT Positions tab
If Range("A5") > 0 Then
    With ActiveSheet
        Set rangecopy = .Range(.Range("A5"), .Cells(5, Columns.Count).End(xlToLeft))
        Set rangepaste = .Range(.Range("A5"), .Cells(Rows.Count, 1).End(xlUp)).Resize(, rangecopy.Columns.Count)
    End With
    rangecopy.Copy
    rangepaste.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    Range("A1").Select
End If


ActiveWorkbook.Worksheets("Multiple YT All Data").Activate 'copying formats for Multiple YT All Data tab
If Range("A5") > 0 Then
    With ActiveSheet
        Set rangecopy = .Range(.Range("A5"), .Cells(5, Columns.Count).End(xlToLeft))
        Set rangepaste = .Range(.Range("A5"), .Cells(Rows.Count, 1).End(xlUp)).Resize(, rangecopy.Columns.Count)
    End With
    rangecopy.Copy
    rangepaste.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    Range("A1").Select
End If


ActiveWorkbook.Worksheets("After Deadline").Activate 'copying formats for After Deadline tab
If Range("A5") > 0 Then
    With ActiveSheet
        Set rangecopy = .Range(.Range("A5"), .Cells(5, Columns.Count).End(xlToLeft))
        Set rangepaste = .Range(.Range("A5"), .Cells(Rows.Count, 1).End(xlUp)).Resize(, rangecopy.Columns.Count)
    End With
    rangecopy.Copy
    rangepaste.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    Range("A1").Select
End If


ActiveWorkbook.Worksheets("DVF after QAD & MDA").Activate 'copying formats for DVF after QAD & MDA tab
If Range("A5") > 0 Then
    With ActiveSheet
        Set rangecopy = .Range(.Range("A5"), .Cells(5, Columns.Count).End(xlToLeft))
        Set rangepaste = .Range(.Range("A5"), .Cells(Rows.Count, 1).End(xlUp)).Resize(, rangecopy.Columns.Count)
    End With
    rangecopy.Copy
    rangepaste.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    Range("A1").Select
End If


End Sub
