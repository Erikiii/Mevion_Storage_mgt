Attribute VB_Name = "Module1"
Option Explicit

' ==========================================
' Module: StorageManagement
' Description: VBA macros for Storage Management System
' ==========================================

' ------------------------------------------
' Macro 1: Record Inbound
' Records the Data Entry form data to Database as Inbound
' ------------------------------------------
Sub RecordInbound()
    Dim wsEntry As Worksheet
    Dim wsDB As Worksheet
    Dim nextRow As Long
    
    Set wsEntry = ThisWorkbook.Sheets("Data Entry")
    Set wsDB = ThisWorkbook.Sheets("Database")
    
    ' Validate input
    If wsEntry.Range("B4").Value = "" Then
        MsgBox "Please enter a P/N", vbExclamation, "Validation Error"
        wsEntry.Range("B4").Select
        Exit Sub
    End If
    
    If wsEntry.Range("B5").Value = "" Then
        MsgBox "Please enter a Product Name", vbExclamation, "Validation Error"
        wsEntry.Range("B5").Select
        Exit Sub
    End If
    
    If wsEntry.Range("B7").Value = "" Or Not IsNumeric(wsEntry.Range("B7").Value) Then
        MsgBox "Please enter a valid Quantity (number)", vbExclamation, "Validation Error"
        wsEntry.Range("B7").Select
        Exit Sub
    End If
    
    ' Find next empty row in Database
    nextRow = wsDB.Cells(wsDB.Rows.Count, "A").End(xlUp).Row + 1
    
    ' Write data to Database
    wsDB.Cells(nextRow, 1).Value = wsEntry.Range("B4").Value  ' P/N
    wsDB.Cells(nextRow, 2).Value = wsEntry.Range("B5").Value  ' Product Name
    wsDB.Cells(nextRow, 3).Value = wsEntry.Range("B6").Value  ' S/N
    wsDB.Cells(nextRow, 4).Value = wsEntry.Range("B7").Value  ' Quantity
    wsDB.Cells(nextRow, 5).Value = wsEntry.Range("B8").Value  ' Unit
    wsDB.Cells(nextRow, 6).Value = wsEntry.Range("B9").Value  ' Unit Price
    wsDB.Cells(nextRow, 7).Value = wsEntry.Range("B7").Value * wsEntry.Range("B9").Value  ' Total Value
    wsDB.Cells(nextRow, 8).Value = wsEntry.Range("B11").Value  ' Date
    wsDB.Cells(nextRow, 9).Value = "Inbound"  ' Status
    wsDB.Cells(nextRow, 10).Value = wsEntry.Range("B13").Value  ' Notes
    
    ' Update Dashboard
    Application.Run "UpdateDashboard"
    
    ' Clear form
    wsEntry.Range("B4:B13").ClearContents
    'wsEntry.Range("B6").Value = "SN" & Format(nextRow, "000")
    wsEntry.Range("B10").Value = "=B7*B9"
    wsEntry.Range("B11").Value = "=TODAY()"
    wsEntry.Range("B12").Value = "Inbound"
    
    MsgBox "Inbound record added successfully!" & vbCrLf & vbCrLf & _
           "P/N: " & wsDB.Cells(nextRow, 1).Value & vbCrLf & _
           "Product: " & wsDB.Cells(nextRow, 2).Value & vbCrLf & _
           "Quantity: " & wsDB.Cells(nextRow, 4).Value, _
           vbInformation, "Record Added"
    
    wsDB.Rows(nextRow).Interior.Color = RGB(226, 239, 218) ' Light green highlight
    Application.Wait Now + TimeValue("0:00:01")
    wsDB.Rows(nextRow).Interior.Color = xlNone
End Sub

' ------------------------------------------
' Macro 2: Record Outbound
' Records the Data Entry form data to Database as Outbound
' ------------------------------------------
Sub RecordOutbound()
    Dim wsEntry As Worksheet
    Dim wsDB As Worksheet
    Dim nextRow As Long
    
    Set wsEntry = ThisWorkbook.Sheets("Data Entry")
    Set wsDB = ThisWorkbook.Sheets("Database")
    
    ' Validate input
    If wsEntry.Range("B4").Value = "" Then
        MsgBox "Please enter a P/N", vbExclamation, "Validation Error"
        wsEntry.Range("B4").Select
        Exit Sub
    End If
    
    If wsEntry.Range("B5").Value = "" Then
        MsgBox "Please enter a Product Name", vbExclamation, "Validation Error"
        wsEntry.Range("B5").Select
        Exit Sub
    End If
    
    If wsEntry.Range("B7").Value = "" Or Not IsNumeric(wsEntry.Range("B7").Value) Then
        MsgBox "Please enter a valid Quantity (number)", vbExclamation, "Validation Error"
        wsEntry.Range("B7").Select
        Exit Sub
    End If
    
    ' Find next empty row in Database
    nextRow = wsDB.Cells(wsDB.Rows.Count, "A").End(xlUp).Row + 1
    
    ' Write data to Database
    wsDB.Cells(nextRow, 1).Value = wsEntry.Range("B4").Value  ' P/N
    wsDB.Cells(nextRow, 2).Value = wsEntry.Range("B5").Value  ' Product Name
    wsDB.Cells(nextRow, 3).Value = wsEntry.Range("B6").Value  ' S/N
    wsDB.Cells(nextRow, 4).Value = wsEntry.Range("B7").Value  ' Quantity
    wsDB.Cells(nextRow, 5).Value = wsEntry.Range("B8").Value  ' Unit
    wsDB.Cells(nextRow, 6).Value = wsEntry.Range("B9").Value  ' Unit Price
    wsDB.Cells(nextRow, 7).Value = wsEntry.Range("B7").Value * wsEntry.Range("B9").Value  ' Total Value
    wsDB.Cells(nextRow, 8).Value = wsEntry.Range("B11").Value  ' Date
    wsDB.Cells(nextRow, 9).Value = "Outbound"  ' Status
    wsDB.Cells(nextRow, 10).Value = wsEntry.Range("B13").Value  ' Notes
    
    ' Update Dashboard
    Application.Run "UpdateDashboard"
    
    ' Clear form
    wsEntry.Range("B4:B13").ClearContents
    'wsEntry.Range("B4").Value = "SN" & Format(nextRow, "000")
    wsEntry.Range("B12").Value = "Outbound"
    wsEntry.Range("B11").Value = Date
    
    MsgBox "Outbound record added successfully!" & vbCrLf & vbCrLf & _
           "P/N: " & wsDB.Cells(nextRow, 1).Value & vbCrLf & _
           "Product: " & wsDB.Cells(nextRow, 2).Value & vbCrLf & _
           "Quantity: " & wsDB.Cells(nextRow, 4).Value, _
           vbInformation, "Record Added"
    
    wsDB.Rows(nextRow).Interior.Color = RGB(252, 228, 236) ' Light red highlight
    Application.Wait Now + TimeValue("0:00:01")
    wsDB.Rows(nextRow).Interior.Color = xlNone
End Sub

' ------------------------------------------
' Macro 3: Search Records
' Searches Database based on criteria and displays results
' ------------------------------------------
Sub SearchRecords()
    Dim wsSearch As Worksheet
    Dim wsDB As Worksheet
    Dim searchSN As String
    Dim searchKeyword As String
    Dim searchStatus As String
    Dim lastRow As Long
    Dim resultRow As Long
    Dim found As Boolean
    
    Set wsSearch = ThisWorkbook.Sheets("Search")
    Set wsDB = ThisWorkbook.Sheets("Database")
    
    ' Get search criteria
    searchSN = Trim(wsSearch.Range("B4").Value)
    searchKeyword = Trim(wsSearch.Range("B5").Value)
    searchStatus = Trim(wsSearch.Range("B6").Value)
    
    ' Clear previous results
    If wsDB.Rows.Count > 100 Then
        wsSearch.Range("A9:J1000").ClearContents
    End If
    
    ' Find last row in Database
    lastRow = wsDB.Cells(wsDB.Rows.Count, "A").End(xlUp).Row
    
    If lastRow < 2 Then
        MsgBox "No records found in Database", vbExclamation, "Search"
        Exit Sub
    End If
    
    ' Search and display results
    resultRow = 9
    found = False
    
    Dim i As Long
    For i = 2 To lastRow
        ' Check P/N match
        Dim snMatch As Boolean
        snMatch = True
        If searchSN <> "" And searchSN <> "0" Then
            If InStr(CStr(wsDB.Cells(i, 1).Value), searchSN) = 0 Then
                snMatch = False
            End If
        End If
        
        ' Check keyword match (Product Name or Category)
        Dim kwMatch As Boolean
        kwMatch = True
        If searchKeyword <> "" Then
            If InStr(LCase(CStr(wsDB.Cells(i, 2).Value)), LCase(searchKeyword)) = 0 And _
            InStr(LCase(CStr(wsDB.Cells(i, 3).Value)), LCase(searchKeyword)) = 0 Then
                kwMatch = False
            End If
        End If
        
        ' Check status match
        Dim statusMatch As Boolean
        statusMatch = True
        If searchStatus <> "All" And searchStatus <> "" Then
            If CStr(wsDB.Cells(i, 9).Value) <> searchStatus Then
                statusMatch = False
            End If
        End If
        
        ' If all criteria match, display the row
        If snMatch And kwMatch And statusMatch Then
            found = True
            Dim j As Long
            For j = 1 To 10
                wsSearch.Cells(resultRow, j).Value = wsDB.Cells(i, j).Value
            Next j
            resultRow = resultRow + 1
        End If
    Next i
    
    ' Apply formatting to results
    If found Then
        ' Add borders and formatting
        Dim r As Long
        For r = 9 To resultRow - 1
            Dim c As Long
            For c = 1 To 10
                With wsSearch.Cells(r, c)
                    .Font.Name = "Calibri"
                    .Font.Size = 10
                    .HorizontalAlignment = xlCenter
                    .Borders(xlEdgeLeft).LineStyle = xlContinuous
                    .Borders(xlEdgeRight).LineStyle = xlContinuous
                    .Borders(xlEdgeTop).LineStyle = xlContinuous
                    .Borders(xlEdgeBottom).LineStyle = xlContinuous
                End With
            Next c
            ' Alternate row colors
            If (r - 9) Mod 2 = 1 Then
                For c = 1 To 10
                    wsSearch.Cells(r, c).Interior.Color = RGB(242, 242, 242)
                Next c
            End If
        Next r
        
        MsgBox "Search complete! Found " & (resultRow - 9) & " record(s).", vbInformation, "Search Results"
    Else
        MsgBox "No records match your search criteria.", vbExclamation, "Search"
    End If
End Sub

' ------------------------------------------
' Macro 4: Update Dashboard
' Refreshes the Dashboard calculations
' ------------------------------------------
Sub UpdateDashboard()
    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    
    ' Force recalculation
    Application.CalculateFullRebuild
    
    MsgBox "Dashboard updated successfully!", vbInformation, "Update"
End Sub

' ------------------------------------------
' Macro 5: Add Button to Data Entry Sheet
' Creates Inbound and Outbound buttons
' ------------------------------------------
Sub AddButtons()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Data Entry")
    
    ' Add Inbound button
    With ws.Buttons.Add(280, 40, 120, 30)
        .OnAction = "RecordInbound"
        .Caption = "Record Inbound"
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        '.Interior.Color = RGB(89, 161, 79) ' Green
        .Font.Color = RGB(255, 255, 255)
    End With
    
    ' Add Outbound button
    With ws.Buttons.Add(280, 80, 120, 30)
        .OnAction = "RecordOutbound"
        .Caption = "Record Outbound"
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        '.Interior.Color = RGB(225, 87, 89) ' Red
        '.Font.Color = RGB(255, 255, 255)
    End With
    
    ' Add Search button
    Dim wsSearch As Worksheet
    Set wsSearch = ThisWorkbook.Sheets("Search")
    With wsSearch.Buttons.Add(280, 40, 120, 30)
        .OnAction = "SearchRecords"
        .Caption = "Search"
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        '.Interior.Color = RGB(46, 117, 182) ' Blue
        '.Font.Color = RGB(255, 255, 255)
    End With
    
    MsgBox "Buttons added successfully!" & vbCrLf & vbCrLf & _
           "Data Entry sheet: Record Inbound / Record Outbound buttons" & vbCrLf & _
           "Search sheet: Search button", _
           vbInformation, "Buttons Added"
End Sub

' ------------------------------------------
' Macro 6: Refresh Worksheet
' Refreshes Worksheet to have correct date
' ------------------------------------------
Sub RefreshOnOpening()
    Dim wsEntry As Worksheet
    
    Set wsEntry = ThisWorkbook.Sheets("Data Entry")
    
    ' Clear form
    wsEntry.Range("B4:B13").ClearContents
    wsEntry.Range("B10").Value = "=B7*B9"
    wsEntry.Range("B11").Value = "=TODAY()"
    
    MsgBox "Sheet refreshed successfully!"
End Sub


