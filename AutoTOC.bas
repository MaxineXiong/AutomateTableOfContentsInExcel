Sub TOC_Generator()
'This macro generates a Table of Contents for an Excel workbook.
    Dim sh As Worksheet
    Dim selectedRange As Range
    Dim startTOC As Range
    Dim endTOC As Range
    Dim currentCell As Range
    Dim inNewSheet As VbMsgBoxResult
    Dim numVisibleSheets As Byte
    Dim ToOverwrite As VbMsgBoxResult
    
    'Disable display alerts to avoid interruption by alert messages.
    Application.DisplayAlerts = False

    'Handle errors gracefully using error handling
    On Error GoTo ErrorHandler
    'Prompt the user if they want to place the Table of Contents in a new sheet.
    inNewSheet = MsgBox("Would you like to place the Table of Content in a new sheet?", vbYesNoCancel, "New Sheet for Table of Content?")
    
    'If the user selects "Yes"...
    If inNewSheet = vbYes Then
        'Add a new sheet and name it "TOC".
        Sheets.Add Sheets(1)
        Sheets(1).Name = "TOC"
    'If the user selects "Cancel"...
    ElseIf inNewSheet = vbCancel Then
        'Exit the program
        Exit Sub
    End If
    
    'Prompt the user to select a starting point for the Table of Contents.
    Set selectedRange = Excel.Application.InputBox("Where would you like to place the Table of Contents within this sheet?" _
                        & vbNewLine & "Please choose a starting point, such as a specific cell:", "Insert Table of Content", , , , , , 8)
    
    'Count the number of visible sheets in the workbook.
    numVisibleSheets = 0
    For Each sh In ThisWorkbook.Sheets
        If sh.Visible = True Then
            numVisibleSheets = numVisibleSheets + 1
        End If
    Next sh

    'Define the start and end points for the Table of Contents, based on the selected range and the number of visible sheets
    Set startTOC = selectedRange.Cells(1, 1)
    Set endTOC = startTOC.Offset(numVisibleSheets - 2, 1)
    Set currentCell = startTOC
    
    'Check if the range to accomodate Table of Contents contains any values.
    'If the range contains values...
    If Excel.WorksheetFunction.CountA(Range(startTOC, endTOC)) > 0 Then
        'Prompt the user if they want to overwrite existing content within the TOC range.
        ToOverwrite = MsgBox("The values in the range " & Replace(startTOC.Address, "$", "") _
                             & ":" & Replace(endTOC.Address, "$", "") & " will be overwritten." _
                             & vbNewLine & "Would you like to continue?", _
                             vbOKCancel + vbDefaultButton2, "Overwritting existing content?")
        If ToOverwrite = vbOK Then
            'If confirmed, proceed to createTOC label
            GoTo createTOC
        Else
            'If not confirmed, exit the program
            Exit Sub
        End If
    
    'If the range is empty...
    Else
        ' Generate the Table of Contents by looping through each visible sheet in this workbook.
createTOC:         For Each sh In ThisWorkbook.Sheets
                        If sh.Visible = True And sh.Name <> ActiveSheet.Name Then
                            'Create hyperlinks to each sheet and display their sheet names in the first column of TOC.
                            ActiveSheet.Hyperlinks.Add currentCell, Address:="", _
                                SubAddress:="'" & sh.Name & "'!A1", TextToDisplay:=sh.Name
                            'Display the topic covered in the second column of TOC.
                            currentCell.Offset(0, 1).Value = sh.Range("A1").Value
                            'Move to the next cell for the next sheet
                            Set currentCell = currentCell.Offset(1, 0)
                        End If
                   Next sh
    End If
    
    'Format the Table of Contents range.
    'Align the content to the left and center vertically. Set the font size as 11
    With ActiveSheet.Range(startTOC, endTOC)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Size = 11
    End With
    'Make the second column of TOC bold
    ActiveSheet.Range(endTOC.End(xlUp), endTOC).Font.Bold = True
    'Autofit rows height and columns width in the TOC sheet to optimize readability
    ActiveSheet.Rows.AutoFit
    ActiveSheet.Columns.AutoFit

    'Re-enable display alerts
    Application.DisplayAlerts = True

Exit Sub
ErrorHandler:
    'Handle specific errors encountered during the execution of the macro.
    Select Case Err.Number
        Case 1004
            'If the "TOC" sheet already exists, prompt the user for further action.
            Sheets("TOC").Select
            If MsgBox("The sheet ""TOC"" already exists. Would you like to place the Table of Content in this sheet?", _
            vbExclamation + vbYesNo, "TOC Sheet already exists") = vbYes Then
                'Delete the newly added sheet and proceed with generating the TOC in the "TOC" sheet.
                Sheets(1).Delete
            Else:
                'Rename the newly added sheet to a unique name based on the number of sheets in this workbook.
                Sheets(1).Name = "Sheets" & ThisWorkbook.Sheets.Count
                Sheets(1).Select
            End If
            'Proceed with generating the TOC in the current sheet
            Resume Next
        Case 424
            'If no start point is selected for the TOC, exit the program
            Exit Sub
        Case Else
            'Display a generic error message for other encountered errors.
            MsgBox "An error has occurred!", vbExclamation, "Error Found"
    End Select
    
End Sub



