Attribute VB_Name = "Module1"
Sub A1StartAreaUPV()
    Dim wsRecieveUPV As Worksheet
    Dim wsDeLimiter As Worksheet
    Dim wsAreaCompletion As Worksheet

    ' Set the worksheets
    On Error Resume Next
    Set wsRecieveUPV = ThisWorkbook.Sheets("RecieveUPVFileHere")
    Set wsDeLimiter = ThisWorkbook.Sheets("DeLimiter")
    Set wsAreaCompletion = ThisWorkbook.Sheets("AreaCompletionTable")
    On Error GoTo 0

    ' Ensure sheets exist before clearing them
    If Not wsRecieveUPV Is Nothing Then wsRecieveUPV.Cells.Clear
    If Not wsDeLimiter Is Nothing Then wsDeLimiter.Cells.Clear
    If Not wsAreaCompletion Is Nothing Then wsAreaCompletion.Cells.Clear

    ' Call AreaCode subroutine
    Call AreaCode

    'MsgBox "Close this message box and wait for the next one.", vbInformation
End Sub



Sub AreaCode()
    Dim wsUPV As Worksheet
    Dim wsTemplate As Worksheet
    Dim AreaCode As Long
    Dim AreaCol As Long
    Dim lastRow As Long
    Dim i As Long
    
    ' Set the worksheets
    Set wsUPV = ThisWorkbook.Sheets("Input your area here")
    Set wsTemplate = ThisWorkbook.Sheets("Template")
    
    ' Get the AreaCode from F2 in UPV sheet
    AreaCode = wsUPV.Range("F2").Value
    
    ' Debugging: Output the AreaCode value to the Immediate Window
    Debug.Print "AreaCode: " & AreaCode
    
    ' Find the Area column in Template sheet (row 8)
    AreaCol = 0 ' Initialize to 0
    For i = 1 To wsTemplate.Cells(8, wsTemplate.Columns.Count).End(xlToLeft).Column
        If wsTemplate.Cells(8, i).Value = "Area" Then
            AreaCol = i
            Exit For
        End If
    Next i
    
    If AreaCol = 0 Then
        MsgBox "Area column not found!", vbExclamation
        Exit Sub
    End If
    
    ' Find the last row in the AreaCol
    lastRow = wsTemplate.Cells(wsTemplate.Rows.Count, AreaCol).End(xlUp).Row
    
    ' Loop from the bottom to row 10 and delete rows where AreaCol value does not match AreaCode
    For i = lastRow To 10 Step -1
        If wsTemplate.Cells(i, AreaCol).Value <> AreaCode Then
            Debug.Print "Deleting row " & i & " with Area value: " & wsTemplate.Cells(i, AreaCol).Value ' Debugging: Output the row and value being deleted
            wsTemplate.Rows(i).Delete
        End If
    Next i
    'MsgBox "Rows filtered and deleted based on AreaCode", vbInformation
    Call AreaLightingFilter
End Sub

Sub AreaLightingFilter()
    Dim wsUPV As Worksheet
    Dim wsTemplate As Worksheet
    Dim AreaCode As Long
    Dim AreaCol As Long
    Dim NameCol As Long
    Dim QuantityCol As Long
    Dim TDQuantityCol As Long
    Dim CompleteCol As Long
    Dim lastRow As Long
    Dim i As Long
    
    ' Set the worksheets
    Set wsUPV = ThisWorkbook.Sheets("Input your area here")
    Set wsTemplate = ThisWorkbook.Sheets("Template")
    
    ' Get the AreaCode from F2 in UPV sheet
    AreaCode = wsUPV.Range("F2").Value
    
    ' Debugging: Output the AreaCode value to the Immediate Window
    Debug.Print "AreaCode: " & AreaCode
    
    ' Find the Area column in Template sheet (row 8)
    AreaCol = 0 ' Initialize to 0
    For i = 1 To wsTemplate.Cells(8, wsTemplate.Columns.Count).End(xlToLeft).Column
        If wsTemplate.Cells(8, i).Value = "Area" Then
            AreaCol = i
            Exit For
        End If
    Next i
    
    If AreaCol = 0 Then
        MsgBox "Area column not found!", vbExclamation
        Exit Sub
    End If
    
    ' Find the other required columns in Template sheet (row 8)
    NameCol = 0
    QuantityCol = 0
    TDQuantityCol = 0
    CompleteCol = 0
    
    For i = 1 To wsTemplate.Cells(8, wsTemplate.Columns.Count).End(xlToLeft).Column
        Select Case wsTemplate.Cells(8, i).Value
            Case "Name"
                NameCol = i
            Case "Quantity"
                QuantityCol = i
            Case "To Date Quantity"
                TDQuantityCol = i
            Case "% Complete"
                CompleteCol = i
        End Select
    Next i
    
    ' Check if all columns are found
    If NameCol = 0 Or QuantityCol = 0 Or TDQuantityCol = 0 Or CompleteCol = 0 Then
        MsgBox "One or more required columns not found!", vbExclamation
        Exit Sub
    End If
    
    ' Find the new last row in AreaCol (after deletions)
    lastRow = wsTemplate.Cells(wsTemplate.Rows.Count, AreaCol).End(xlUp).Row
    
    ' Loop from the bottom to row 10 and delete rows where NameCol contains "ELP" or "LPP"
    For i = lastRow To 10 Step -1
        If InStr(1, wsTemplate.Cells(i, NameCol).Value, "ELP") > 0 Or InStr(1, wsTemplate.Cells(i, NameCol).Value, "LPP") > 0 Then
            Debug.Print "Deleting row " & i & " with Name value: " & wsTemplate.Cells(i, NameCol).Value ' Debugging: Log the row being deleted
            wsTemplate.Rows(i).Delete
        End If
    Next i
    
    'MsgBox "Rows filtered and deleted based on AreaCode and Name column", vbInformation
    Call AreaReFormatting
End Sub


Sub AreaReFormatting()
    Dim wsUPV As Worksheet
    Dim wsTemplate As Worksheet
    Dim AreaCode As Long
    Dim AreaCol As Long
    Dim NameCol As Long
    Dim lastRow As Long
    Dim i As Long
    Dim originalValue As String
    Dim cablePullPos As Integer
    
    ' Set the worksheets
    Set wsUPV = ThisWorkbook.Sheets("Input your area here")
    Set wsTemplate = ThisWorkbook.Sheets("Template")
    
    ' Get the AreaCode from F2 in UPV sheet
    AreaCode = wsUPV.Range("F2").Value
    
    ' Debugging: Output the AreaCode value to the Immediate Window
    Debug.Print "AreaCode: " & AreaCode
    
    ' Find the Area column in Template sheet (row 8)
    AreaCol = 0 ' Initialize to 0
    For i = 1 To wsTemplate.Cells(8, wsTemplate.Columns.Count).End(xlToLeft).Column
        If wsTemplate.Cells(8, i).Value = "Area" Then
            AreaCol = i
            Exit For
        End If
    Next i
    
    If AreaCol = 0 Then
        MsgBox "Area column not found!", vbExclamation
        Exit Sub
    End If

    ' Find the Name column in Template sheet (row 8)
    NameCol = 0
    For i = 1 To wsTemplate.Cells(8, wsTemplate.Columns.Count).End(xlToLeft).Column
        If wsTemplate.Cells(8, i).Value = "Name" Then
            NameCol = i
            Exit For
        End If
    Next i

    If NameCol = 0 Then
        MsgBox "Name column not found!", vbExclamation
        Exit Sub
    End If

    ' Find the last used row in AreaCol
    lastRow = wsTemplate.Cells(wsTemplate.Rows.Count, AreaCol).End(xlUp).Row

    ' Loop from the bottom to row 10 and delete rows where NameCol contains "ELP", "LPP", or "RPP"
    For i = lastRow To 10 Step -1
        If InStr(1, wsTemplate.Cells(i, NameCol).Value, "ELP") > 0 Or _
           InStr(1, wsTemplate.Cells(i, NameCol).Value, "LPP") > 0 Or _
           InStr(1, wsTemplate.Cells(i, NameCol).Value, "RPP") > 0 Then
            
            Debug.Print "Deleting row " & i & " with Name value: " & wsTemplate.Cells(i, NameCol).Value
            wsTemplate.Rows(i).Delete
        End If
    Next i
    
    ' Loop again to remove everything after " - Cable Pull" in NameCol
    For i = lastRow To 10 Step -1
        originalValue = wsTemplate.Cells(i, NameCol).Value
        cablePullPos = InStr(1, originalValue, " - Cable Pull")

        If cablePullPos > 0 Then
            ' Keep only the part before " - Cable Pull"
            wsTemplate.Cells(i, NameCol).Value = Left(originalValue, cablePullPos - 1)
            
            ' Debugging output
            Debug.Print "Updated Name value at row " & i & ": " & wsTemplate.Cells(i, NameCol).Value
        End If
    Next i
    
    'MsgBox "Rows filtered, deleted, and Name column cleaned.", vbInformation
    Call AreaValidateAndCopy
End Sub




Sub AreaValidateAndCopy()
    Dim wsTemplate As Worksheet
    Dim wsAreaCompletion As Worksheet
    Dim wsCableSchedule As Worksheet
    Dim wsInputArea As Worksheet
    Dim MCSColD As Range
    Dim MCSColQ As Range
    Dim NameCol As Long
    Dim CompleteCol As Long
    Dim lastRowTemplate As Long
    Dim lastPasteRow As Long
    Dim i As Long
    Dim foundCell As Range
    Dim nameValue As String
    Dim mcsValue As Variant
    Dim AreaCol As Long
    Dim highlightedNames As String
    Dim completeValue As Variant
    Dim filterType As String
    
    highlightedNames = "" ' Initialize a string to hold highlighted NameCol values
    
    ' Set the worksheets
    Set wsTemplate = ThisWorkbook.Sheets("Template")
    Set wsAreaCompletion = ThisWorkbook.Sheets("AreaCompletionTable")
    Set wsCableSchedule = ThisWorkbook.Sheets("CableScheduleData")
    Set wsInputArea = ThisWorkbook.Sheets("Input your area here") ' Sheet with G2 value
    
    ' Get filter type from G2 in "Input your area here"
    filterType = Trim(wsInputArea.Range("G2").Value)
    
    ' Get the Area column in Template sheet (row 8)
    AreaCol = 0 ' Initialize to 0
    For i = 1 To wsTemplate.Cells(8, wsTemplate.Columns.Count).End(xlToLeft).Column
        If wsTemplate.Cells(8, i).Value = "Area" Then
            AreaCol = i
            Exit For
        End If
    Next i
    
    If AreaCol = 0 Then
        MsgBox "Area column not found!", vbExclamation
        Exit Sub
    End If

    ' Find the Name and Complete columns in Template sheet (row 8)
    NameCol = 0
    CompleteCol = 0
    
    For i = 1 To wsTemplate.Cells(8, wsTemplate.Columns.Count).End(xlToLeft).Column
        Select Case wsTemplate.Cells(8, i).Value
            Case "Name"
                NameCol = i
            Case "% Complete"
                CompleteCol = i
        End Select
    Next i
    
    ' Check if columns are found
    If NameCol = 0 Or CompleteCol = 0 Then
        MsgBox "One or more required columns not found!", vbExclamation
        Exit Sub
    End If
    
    ' Get the columns for MCSColD and MCSColQ
    Set MCSColD = wsCableSchedule.Range("D1:D" & wsCableSchedule.Cells(wsCableSchedule.Rows.Count, "D").End(xlUp).Row)
    Set MCSColQ = wsCableSchedule.Range("Q1:Q" & wsCableSchedule.Cells(wsCableSchedule.Rows.Count, "Q").End(xlUp).Row)
    
    ' Get the last used row in "AreaCompletionTable" to determine where to paste
    lastPasteRow = wsAreaCompletion.Cells(wsAreaCompletion.Rows.Count, 2).End(xlUp).Row + 1

    ' Get the last row of NameCol (from Template sheet)
    lastRowTemplate = wsTemplate.Cells(wsTemplate.Rows.Count, NameCol).End(xlUp).Row
    
    ' Loop through the rows in NameCol
    For i = 10 To lastRowTemplate
        ' Get the NameCol value
        nameValue = wsTemplate.Cells(i, NameCol).Value
        
        ' Check for errors or blank values in CompleteCol
        If IsError(wsTemplate.Cells(i, CompleteCol).Value) Or IsEmpty(wsTemplate.Cells(i, CompleteCol).Value) Then
            completeValue = 0 ' Default to 0 if there's an error or blank
        Else
            completeValue = wsTemplate.Cells(i, CompleteCol).Value
        End If
        
        ' Debugging: Print row data
        Debug.Print "Row: " & i & " | Name: " & nameValue & " | Complete: " & completeValue
        
        ' Apply filter conditions based on G2 value
        Select Case filterType
            Case "All"
                ' Do not filter�allow all values through
                
            Case "Partial"
                If completeValue <= 0 Or completeValue >= 95 Then GoTo SkipRow ' Only allow 0 < x < 95
                
            Case "Complete"
                If completeValue < 95 Or completeValue > 100 Then GoTo SkipRow ' Only allow 95 = x = 100
                
            Case "Zeroed"
                If completeValue <> 0 Then GoTo SkipRow ' Only allow x = 0
                
            Case Else
                ' If G2 has an unexpected value, exit to prevent errors
                MsgBox "Invalid filter type in G2: " & filterType, vbExclamation
                Exit Sub
        End Select
        
        ' Look for the nameValue in MCSColD
        Set foundCell = MCSColD.Find(nameValue, LookIn:=xlValues, LookAt:=xlWhole)
        
        If foundCell Is Nothing Then
            ' If the value is not found, highlight the cell in Template sheet
            wsTemplate.Cells(i, NameCol).Interior.Color = RGB(255, 0, 0) ' Red highlight
            
            ' Add the Name value to the string of highlighted names
            highlightedNames = highlightedNames & nameValue & vbCrLf
        Else
            ' If the value is found, copy NameCol value to AreaCompletionTable in the next available row
            wsAreaCompletion.Cells(lastPasteRow, 2).Value = nameValue ' Column B
            
            ' Copy the associated MCSColQ value into Column C of AreaCompletionTable
            mcsValue = MCSColQ.Cells(foundCell.Row - MCSColD.Row + 1, 1).Value
            wsAreaCompletion.Cells(lastPasteRow, 3).Value = mcsValue ' Column C
            
            ' Move to the next row for pasting
            lastPasteRow = lastPasteRow + 1
        End If
        
SkipRow:
    Next i
    
    ' If any rows were highlighted, show a message box
    If highlightedNames <> "" Then
        MsgBox "The following NameCol values were not found in MCSColD and have been highlighted:" & vbCrLf & highlightedNames, vbExclamation
    End If
    
    'MsgBox "Step 4 complete: Validation and Data Copying finished.", vbInformation
    Call AreaDelimitAndTranspose
End Sub



Sub AreaDelimitAndTranspose()
    Dim wsAreaCompletion As Worksheet
    Dim wsDeLimiter As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cblName As String
    Dim cblRouting As String
    Dim routingArray As Variant
    Dim headerCol As Long
    Dim rowIdx As Long
    
    ' Set the worksheets
    Set wsAreaCompletion = ThisWorkbook.Sheets("AreaCompletionTable")
    Set wsDeLimiter = ThisWorkbook.Sheets("DeLimiter")
    
    ' Clear previous data in DeLimiter sheet (except for headers)
    wsDeLimiter.Cells.ClearContents
    
    ' Get the last row in AreaCompletionTable
    lastRow = wsAreaCompletion.Cells(wsAreaCompletion.Rows.Count, 2).End(xlUp).Row ' Column B
    
    ' Loop through each row in AreaCompletionTable, starting from row 2
    For i = 2 To lastRow
        ' Check if column C contains an error (#N/A)
        If Not IsError(wsAreaCompletion.Cells(i, 3).Value) Then
            ' Get the CblNames from AreaCompletionTable
            cblName = Trim(wsAreaCompletion.Cells(i, 2).Value) ' Column B (Trim spaces)
            
            ' Skip empty names
            If cblName = "" Then GoTo NextRow
            
            ' Get the CblRouting from AreaCompletionTable
            cblRouting = Trim(CStr(wsAreaCompletion.Cells(i, 3).Value)) ' Column C (force to string)
            
            ' Remove wrap text formatting (line breaks) from CblRouting
            cblRouting = Replace(cblRouting, Chr(10), " ") ' Replace line breaks (Alt+Enter) with a space
            
            ' Find the header column in DeLimiter corresponding to CblNames
            headerCol = 0
            For rowIdx = 2 To wsDeLimiter.Cells(2, wsDeLimiter.Columns.Count).End(xlToLeft).Column
                If Trim(wsDeLimiter.Cells(2, rowIdx).Value) = cblName Then
                    headerCol = rowIdx
                    Exit For
                End If
            Next rowIdx
            
            ' If header doesn't exist, create it
            If headerCol = 0 Then
                headerCol = wsDeLimiter.Cells(2, wsDeLimiter.Columns.Count).End(xlToLeft).Column + 1
                wsDeLimiter.Cells(2, headerCol).Value = cblName ' Set the header
            End If
            
            ' If CblRouting is not empty, split and transpose it
            If Len(cblRouting) > 0 And cblRouting <> "#N/A" Then
                ' Split the CblRouting value into an array using the comma delimiter
                routingArray = Split(cblRouting, ",")
                
                ' Loop through the routingArray and transpose the values into the correct column in DeLimiter
                For rowIdx = LBound(routingArray) To UBound(routingArray)
                    wsDeLimiter.Cells(rowIdx + 3, headerCol).Value = Trim(routingArray(rowIdx)) ' Trim spaces
                Next rowIdx
            Else
                ' If no valid CblRouting, put "#N/A" under the CblNames column in DeLimiter
                wsDeLimiter.Cells(3, headerCol).Value = "#N/A"
            End If
        Else
            ' Debugging: Print skipped rows
            Debug.Print "Skipping row " & i & " due to #N/A in column C."
        End If
        
NextRow:
    Next i
    
    ' Clear formatting and remove excessive spaces in the range
    With wsDeLimiter.Range("A1:AAA500")
        .ClearFormats ' Clear all formatting
        Dim cell As Range
        For Each cell In .Cells
            If Not IsEmpty(cell.Value) Then
                cell.Value = Trim(Replace(cell.Value, "  ", " ")) ' Remove excessive spaces
            End If
        Next cell
    End With
    
    ' Debugging log to check results
    Debug.Print "Delimitation and transposition complete."
    
    ' Call the next step
    Call AreaFormatUPVPackage
End Sub


Sub AreaFormatUPVPackage()
    Dim wsDeLimiter As Worksheet
    Dim wsRecieveUPV As Worksheet
    Dim lastCol As Long
    Dim lastRow As Long
    Dim lastGRow As Long
    Dim i As Long
    Dim cblName As String
    Dim routingArray As Variant
    Dim packageRow As Long
    Dim rowIdx As Long
    Dim tempArray() As String
    Dim orFlag As Boolean

    ' Set the worksheets
    On Error Resume Next
    Set wsDeLimiter = ThisWorkbook.Sheets("DeLimiter")
    Set wsRecieveUPV = ThisWorkbook.Sheets("RecieveUPVFileHere")
    On Error GoTo 0

    ' Ensure sheets exist
    If wsDeLimiter Is Nothing Or wsRecieveUPV Is Nothing Then
        MsgBox "Error: One or more sheets not found!", vbCritical
        Exit Sub
    End If

    ' Format Column F as text to prevent formula issues
    wsRecieveUPV.Columns(6).NumberFormat = "@"

    ' Get the last column in DeLimiter (Row 2 contains cable names)
    lastCol = wsDeLimiter.Cells(2, wsDeLimiter.Columns.Count).End(xlToLeft).Column

    ' Start writing in RecieveUPV at the last used row in column G
    lastGRow = wsRecieveUPV.Cells(wsRecieveUPV.Rows.Count, 7).End(xlUp).Row
    packageRow = Application.Max(1, lastGRow + 1) ' Ensure it starts at row 1

    ' Loop through each cable in DeLimiter (starting from Column B)
    For i = 2 To lastCol
        ' Get the Cable Name from row 2
        cblName = Trim(wsDeLimiter.Cells(2, i).Value)

        ' Skip empty cable names
        If cblName = "" Then GoTo NextCable

        ' Find the last row in the current column that has routing values
        lastRow = wsDeLimiter.Cells(wsDeLimiter.Rows.Count, i).End(xlUp).Row
        
        ' If lastRow is less than 3, there are no routing values
        If lastRow < 3 Then GoTo NextCable

        ' Check if the first routing value is "error2042"
        If Trim(wsDeLimiter.Cells(3, i).Value) = "error2042" Then
            Debug.Print "Skipping " & cblName & " due to error2042."
            GoTo NextCable
        End If

        ' Convert range to array safely
        If lastRow = 3 Then
            ' Single value case: Store as a single-item array
            ReDim tempArray(0)
            tempArray(0) = wsDeLimiter.Cells(3, i).Value
        Else
            ' Multiple values: Store in a properly sized array
            routingArray = wsDeLimiter.Range(wsDeLimiter.Cells(3, i), wsDeLimiter.Cells(lastRow, i)).Value
            
            ' Convert 2D array into a 1D array
            If IsArray(routingArray) Then
                ReDim tempArray(UBound(routingArray, 1) - 1)
                For rowIdx = 1 To UBound(routingArray, 1)
                    tempArray(rowIdx - 1) = routingArray(rowIdx, 1)
                Next rowIdx
            Else
                ' If it's not an array, manually add the single value
                ReDim tempArray(0)
                tempArray(0) = wsDeLimiter.Cells(3, i).Value
            End If
        End If

        ' Find the last used row in RecieveUPV column G before writing a new package
        lastGRow = wsRecieveUPV.Cells(wsRecieveUPV.Rows.Count, 7).End(xlUp).Row
        packageRow = Application.Max(1, lastGRow + 1)

        ' Debugging: Print package row
        Debug.Print "Writing Package to Row: " & packageRow

        ' Ensure packageRow is valid before writing
        If packageRow <= 0 Or packageRow > 1048576 Then
            Debug.Print "ERROR: packageRow out of valid range!"
            GoTo NextCable
        End If

        ' Set Package header in RecieveUPV
        wsRecieveUPV.Cells(packageRow, 2).Value = cblName ' Column B for Cable Name
        wsRecieveUPV.Cells(packageRow, 1).Value = "Package" ' Column A for "Package"

        ' Move to the next row after the header for the delimited data
        packageRow = packageRow + 1

        ' Reset "Or" flag
        orFlag = False

        ' Loop through each routing value and paste it into RecieveUPV
        For rowIdx = LBound(tempArray) To UBound(tempArray)
            If orFlag Then
                wsRecieveUPV.Cells(packageRow, 3).Value = "Or" ' Column C: Add "Or" for subsequent rows
            End If
            
            ' Debugging: Print packageRow and value
            Debug.Print "Row: " & packageRow & " | Value: " & tempArray(rowIdx)

            ' Ensure packageRow is within a valid range before writing to columns D, E, F, and G
            If packageRow > 0 And packageRow < 1048576 Then
                wsRecieveUPV.Cells(packageRow, 7).Value = tempArray(rowIdx) ' Column G (Routing Value)
                wsRecieveUPV.Cells(packageRow, 6).Value = CStr("==") ' Column F (Now properly formatted as text)
                wsRecieveUPV.Cells(packageRow, 5).Value = "RunName" ' Column E
                wsRecieveUPV.Cells(packageRow, 4).Value = "Attribute" ' Column D
            Else
                Debug.Print "ERROR: packageRow out of bounds!"
                GoTo NextCable
            End If

            ' Move to the next row for the next routing value
            packageRow = packageRow + 1
            orFlag = True ' Set the flag to True after the first row
        Next rowIdx

NextCable:
    Next i

    'MsgBox "Finished, Please check RecieveUPVFileHere", vbInformation
    Call AreaMarkTrayRows
End Sub


Sub AreaMarkTrayRows()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("RecieveUPVFileHere")

    ' Find the last used row in column G
    lastRow = ws.Cells(ws.Rows.Count, 7).End(xlUp).Row ' Column G

    ' Loop through each row in column G
    For i = 1 To lastRow
        cellValue = ws.Cells(i, 7).Value ' Get value from column G

        ' Check if "TRAY" is embedded in the string (case-insensitive)
        If InStr(1, cellValue, "TRAY", vbTextCompare) > 0 Then
            ws.Cells(i, 5).Value = "Name" ' Set column E of that row to "Name"
        End If
    Next i

    MsgBox "Finished, Please check RecieveUPVFileHere", vbInformation
End Sub

Sub A1StartCableScheduleSort()
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim rngVisible As Range
    Dim lastRow As Long

    ' Set the source and target worksheets
    Set wsSource = ThisWorkbook.Sheets("CableScheduleData")
    Set wsTarget = ThisWorkbook.Sheets("CABLE NAME LIST")

    ' Find the last row in column D of the source sheet
    lastRow = wsSource.Cells(wsSource.Rows.Count, "D").End(xlUp).Row

    ' Get the visible (filtered) cells in column D, excluding header (D2 down)
    On Error Resume Next
    Set rngVisible = wsSource.Range("D2:D" & lastRow).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    ' Clear destination range first (optional)
    wsTarget.Range("B2:B10000").ClearContents

    If Not rngVisible Is Nothing Then
        ' Copy visible values into B2 of target sheet
        rngVisible.Copy Destination:=wsTarget.Range("B2")
        Debug.Print "Copied " & rngVisible.Cells.Count & " visible cable names to 'CABLE NAME LIST'!B2"
    Else
        MsgBox "No visible cells in column D to copy.", vbExclamation
    End If
    Call A1StartCableListUPV
End Sub
Sub A1StartCableListUPV()
    Dim wsCableList As Worksheet
    Dim wsCableSchedule As Worksheet
    Dim wsAreaCompletion As Worksheet
    Dim wsDeLimiter As Worksheet
    Dim wsRecieveUPV As Worksheet
    Dim MCSColD As Range
    Dim MCSColQ As Range
    Dim lastRowList As Long
    Dim lastRowSchedule As Long
    Dim lastPasteRow As Long
    Dim i As Long
    Dim cableName As String
    Dim foundCell As Range
    Dim routingValue As Variant
    Dim highlightedNames As String

    ' Set your sheets
    Set wsCableList = ThisWorkbook.Sheets("Cable Name List")
    Set wsCableSchedule = ThisWorkbook.Sheets("CableScheduleData")
    Set wsAreaCompletion = ThisWorkbook.Sheets("AreaCompletionTable")
    Set wsDeLimiter = ThisWorkbook.Sheets("DeLimiter")
    Set wsRecieveUPV = ThisWorkbook.Sheets("RecieveUPVFileHere")

    ' Clear all relevant sheets
    wsAreaCompletion.Cells.ClearContents
    wsDeLimiter.Cells.ClearContents
    wsRecieveUPV.Cells.ClearContents

    ' Set up headers in AreaCompletionTable
    wsAreaCompletion.Range("B1").Value = "Name"
    wsAreaCompletion.Range("C1").Value = "Routing"

    ' Get cable names range
    lastRowList = wsCableList.Cells(wsCableList.Rows.Count, "B").End(xlUp).Row

    ' Define CableScheduleData search ranges
    lastRowSchedule = wsCableSchedule.Cells(wsCableSchedule.Rows.Count, "D").End(xlUp).Row
    Set MCSColD = wsCableSchedule.Range("D1:D" & lastRowSchedule)
    Set MCSColQ = wsCableSchedule.Range("Q1:Q" & lastRowSchedule)

    lastPasteRow = 2
    highlightedNames = ""

    ' Loop through cable names in "Cable Name List"
    For i = 3 To lastRowList
        cableName = Trim(wsCableList.Cells(i, "B").Value)

        If cableName <> "" Then
            Set foundCell = MCSColD.Find(cableName, LookIn:=xlValues, LookAt:=xlWhole)

            If Not foundCell Is Nothing Then
                routingValue = MCSColQ.Cells(foundCell.Row - MCSColD.Row + 1, 1).Value

                ' Paste into AreaCompletionTable
                wsAreaCompletion.Cells(lastPasteRow, "B").Value = cableName
                wsAreaCompletion.Cells(lastPasteRow, "C").Value = routingValue
                lastPasteRow = lastPasteRow + 1
            Else
                ' Not found � highlight and build error list
                wsCableList.Cells(i, "B").Interior.Color = RGB(255, 0, 0)
                highlightedNames = highlightedNames & cableName & vbCrLf
            End If
        End If
    Next i

    ' If any cable names were not found, stop execution and notify
    If highlightedNames <> "" Then
        MsgBox "? Operation cancelled. The following cable names were not found in CableScheduleData and have been highlighted:" & vbCrLf & vbCrLf & highlightedNames, vbCritical
        Exit Sub
    Else
        ' Clear any previous red fills if everything passed
        wsCableList.Range("B3:B" & lastRowList).Interior.ColorIndex = xlNone
    End If

    ' Continue if no issues
    Call AreaDelimitAndTranspose
End Sub
