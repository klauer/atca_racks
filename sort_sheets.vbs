Sub Sort_Active_Book()
Dim i As Integer
Dim j As Integer
Dim iAnswer As VbMsgBoxResult
'
' Prompt the user as which direction they wish to
' sort the worksheets.
'
   iAnswer = MsgBox("Sort Sheets in Ascending Order?" & Chr(10) _
     & "Clicking No will sort in Descending Order", _
     vbYesNoCancel + vbQuestion + vbDefaultButton1, "Sort Worksheets")
   For i = 1 To Sheets.Count
      For j = 1 To Sheets.Count - 1
'
' If the answer is Yes, then sort in ascending order.
'
         If iAnswer = vbYes Then
            If UCase$(Sheets(j).Name) > UCase$(Sheets(j + 1).Name) Then
               Sheets(j).Move After:=Sheets(j + 1)
            End If
'
' If the answer is No, then sort in descending order.
'
         ElseIf iAnswer = vbNo Then
            If UCase$(Sheets(j).Name) < UCase$(Sheets(j + 1).Name) Then
               Sheets(j).Move After:=Sheets(j + 1)
            End If
         End If
      Next j
   Next i
End Sub


Sub PrtSetUp()
Dim ws As Worksheet
Dim rangecell As Range
For Each ws In ThisWorkbook.Worksheets
    Set rangecell = ws.UsedRange.Find("PRINTAREA", MatchCase:=True, LookIn:=xlValues)
    Debug.Print (rangecell.Address)
    print_area = Range("$A$1", rangecell.Offset(-1, -1).Address).Address
    With ws.PageSetup
        .PrintArea = print_area
        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .CenterHorizontally = True
        .CenterVertically = True
        
    End With
    
    ' ws.Name
    Debug.Print (ws.Name)
    'Exit For
Next ws
End Sub
