Sub SplitData()
Dim SplitFld As Range 'the column the end user will select to base the split on
Dim Hdgs As Range 'the headings needed on each worksheet
Dim SplitItem As Range 'the current value in the column that has been selected
Dim NewWs As Worksheet 'a new worksheet as required
Dim ws As Worksheet 'worksheets in the current workbook
Dim WsExists As Boolean 'TRUE or FALSE: does a worksheet already exist for the SplitItem?
Dim SplitWs As Worksheet 'The active worksheet
Set SplitWs = ActiveSheet
On Error GoTo SplitFldError 'if the user cancels the SplitFld inputbox exit sub
'ask user to select the column to base the split on and store that range in the SplitFld variable
Set SplitFld = Application.InputBox _
(Prompt:="Select the column you want to split your data by (***do not include the column heading***)", _
Title:="Column", Type:=8)
On Error GoTo HdgsError 'if the user cancels the Hdgs inputbox exit sub
'ask the user to select the column headings and store that range in the Hdgs variable
Set Hdgs = Application.InputBox _
(Prompt:="Select the headings you want to appear on each worksheet", _
Title:="Headings", Type:=8)
Application.ScreenUpdating = False 'turning off screen updating makes the code run faster
For Each SplitItem In SplitFld 'for each value in the column the user has selected
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = SplitItem Then 'check whether a worksheet already exists for that value
            WsExists = True ' and store TRUE or FALSE in the WsExists variable
            Exit For
        Else
            WsExists = False
        End If
    Next ws
    
    
    If WsExists Then 'if WsExists = TRUE, (if the worksheet does already exist)
    
        'copy the record to the next available row in that worksheet
        Range(SplitItem.End(xlToLeft), SplitItem.End(xlToLeft).End(xlToRight)).Copy _
        Destination:=Worksheets(SplitItem.Value).Range("A1").End(xlDown).Offset(1, 0)
    
        Else 'if WsExists = 'FALSE (if a worksheet doesn't yet exist)
        
        'Create a new worksheet and place it to the right of other worksheets in the workbook
        Set NewWs = Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        'Name the worksheet using the current value stored in the SplitItem variable
        NewWs.Name = SplitItem
        
        'Copy the headings to the new worksheet
        Hdgs.Copy Destination:=NewWs.Range("A1")
        
        'Copy the record to the new worksheet
        Range(SplitItem.End(xlToLeft), SplitItem.End(xlToLeft).End(xlToRight)).Copy Destination:=NewWs.Range("A2")
        
    End If
Next SplitItem
For Each ws In ThisWorkbook.Worksheets 'autofit columns in each worksheet
    ws.UsedRange.Columns.AutoFit
Next ws
'turn screen updating back on
Application.ScreenUpdating = True
Exit Sub
SplitFldError:
Exit Sub
HdgsError:
Exit Sub


Sub AverageColumnsBelowLastRow()
   Dim LastRow As Long
   Dim LastCol As Long
   Dim TargetRow As Long
   Dim TargetRowplus1 As Long
   Dim TotalRows As Long
 
   ' Find the last row in the active sheet
   LastRow = Cells(Rows.Count, 1).End(xlUp).Row
 
   ' Find the last column in the active sheet
   LastCol = Cells(LastRow, Columns.Count).End(xlToLeft).Column
 
   ' Calculate the average for each column and place it in the row below the last row
   TargetRow = LastRow + 1 ' Row just below the last row
   TargetRowplus1 = LastRow + 2 ' 2 rows just below the last row
   TotalRows = LastRow - 1
   
 
   For i = 7 To 35
        Cells(TargetRow, i).Formula = "=COUNTIF(" & Cells(2, i).Address & ":" & Cells(LastRow, i).Address & ",""Yes"") + COUNTIF(" & Cells(2, i).Address & ":" & Cells(LastRow, i).Address & ",""NA"")"
   Next i
   For i = 7 To 35
        Cells(TargetRowplus1, i).Value = Cells(TargetRow, i).Value / TotalRows
   Next i
End Sub
    
End Sub



Sub SplitData()
Dim SplitFld As Range 'the column the end user will select to base the split on
Dim Hdgs As Range 'the headings needed on each worksheet
Dim SplitItem As Range 'the current value in the column that has been selected
Dim newWs As Worksheet 'a new worksheet as required
Dim ws As Worksheet 'worksheets in the current workbook
Dim WsExists As Boolean 'TRUE or FALSE: does a worksheet already exist for the SplitItem?
Dim SplitWs As Worksheet 'The active worksheet
Set SplitWs = ActiveSheet
On Error GoTo SplitFldError 'if the user cancels the SplitFld inputbox exit sub
'ask user to select the column to base the split on and store that range in the SplitFld variable
Set SplitFld = Application.InputBox _
(Prompt:="Select the column you want to split your data by (***do not include the column heading***)", _
Title:="Column", Type:=8)
On Error GoTo HdgsError 'if the user cancels the Hdgs inputbox exit sub
'ask the user to select the column headings and store that range in the Hdgs variable
Set Hdgs = Application.InputBox _
(Prompt:="Select the headings you want to appear on each worksheet", _
Title:="Headings", Type:=8)
Application.ScreenUpdating = False 'turning off screen updating makes the code run faster
For Each SplitItem In SplitFld 'for each value in the column the user has selected
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = SplitItem Then 'check whether a worksheet already exists for that value
            WsExists = True ' and store TRUE or FALSE in the WsExists variable
            Exit For
        Else
            WsExists = False
        End If
    Next ws
    
    
    If WsExists Then 'if WsExists = TRUE, (if the worksheet does already exist)
    
        'copy the record to the next available row in that worksheet
        Range(SplitItem.End(xlToLeft), SplitItem.End(xlToLeft).End(xlToRight)).Copy _
        Destination:=Worksheets(SplitItem.Value).Range("A1").End(xlDown).Offset(1, 0)
    
        Else 'if WsExists = 'FALSE (if a worksheet doesn't yet exist)
        
        'Create a new worksheet and place it to the right of other worksheets in the workbook
        Set newWs = Sheets.Add(after:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        'Name the worksheet using the current value stored in the SplitItem variable
        newWs.Name = SplitItem
        
        'Copy the headings to the new worksheet
        Hdgs.Copy Destination:=newWs.Range("A1")
        
        'Copy the record to the new worksheet
        Range(SplitItem.End(xlToLeft), SplitItem.End(xlToLeft).End(xlToRight)).Copy Destination:=newWs.Range("A2")
        
    End If
Next SplitItem
For Each ws In ThisWorkbook.Worksheets 'autofit columns in each worksheet
    ws.UsedRange.Columns.AutoFit
Next ws
'turn screen updating back on
Application.ScreenUpdating = True
Exit Sub
SplitFldError:
Exit Sub
HdgsError:
Exit Sub
    
End Sub



Sub CountTotalYesNoAndDivideForEachColumn()
   Dim ws As Worksheet
   Set ws = ThisWorkbook.ActiveSheet ' Use the currently active sheet

 

   Dim LastRow As Long
   Dim totalRows As Long
   Dim totalCount As Long
   Dim Col As Long
   Dim Cell As Range

 

   For Col = 2 To 7 ' Columns 2 to 7
       LastRow = ws.Cells(ws.Rows.Count, Col).End(xlUp).Row
       totalRows = LastRow - 1 ' Subtract 1 to exclude the header row
       totalCount = 0 ' Reset count for each column

 

       For Each Cell In ws.Range(ws.Cells(2, Col), ws.Cells(LastRow, Col))
           If Cell.Value = "Yes" Or Cell.Value = "No" Then
               totalCount = totalCount + 1
           End If
       Next Cell

 

       ' Display the total two rows below the last cell in the column
       ws.Cells(LastRow + 2, Col).Value = totalCount

 

       ' Calculate the ratio and display it three rows below the last cell in the column
       If totalRows > 0 Then
           ws.Cells(LastRow + 3, Col).Value = totalCount / totalRows
       Else
           ws.Cells(LastRow + 3, Col).Value = "N/A" ' Avoid division by zero
       End If
   Next Col

 

   For Col = 9 To 11 ' Columns 9 to 11
       LastRow = ws.Cells(ws.Rows.Count, Col).End(xlUp).Row
       totalRows = LastRow - 1 ' Subtract 1 to exclude the header row
       totalCount = 0 ' Reset count for each column

 

       For Each Cell In ws.Range(ws.Cells(2, Col), ws.Cells(LastRow, Col))
           If Cell.Value = "Yes" Or Cell.Value = "No" Then
               totalCount = totalCount + 1
           End If
       Next Cell

 

       ' Display the total two rows below the last cell in the column
       ws.Cells(LastRow + 2, Col).Value = totalCount

 

       ' Calculate the ratio and display it three rows below the last cell in the column
       If totalRows > 0 Then
           ws.Cells(LastRow + 3, Col).Value = totalCount / totalRows
       Else
           ws.Cells(LastRow + 3, Col).Value = "N/A" ' Avoid division by zero
       End If
   Next Col
   
      For Col = 8 To 8 ' Columns 8 to 8
       LastRow = ws.Cells(ws.Rows.Count, Col).End(xlUp).Row
       totalRows = LastRow - 1 ' Subtract 1 to exclude the header row
       totalCount = 0 ' Reset count for each column

 

       For Each Cell In ws.Range(ws.Cells(2, Col), ws.Cells(LastRow, Col))
           If Cell.Value = "Yes" Or Cell.Value = "NA" Then
               totalCount = totalCount + 1
           End If
       Next Cell

 

       ' Display the total two rows below the last cell in the column
       ws.Cells(LastRow + 2, Col).Value = totalCount

 

       ' Calculate the ratio and display it three rows below the last cell in the column
       If totalRows > 0 Then
           ws.Cells(LastRow + 3, Col).Value = totalCount / totalRows
       Else
           ws.Cells(LastRow + 3, Col).Value = "N/A" ' Avoid division by zero
       End If
   Next Col
   
End Sub
