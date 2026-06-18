Sub Merge_PerCase_With_All_Columns()

    Dim wsData As Worksheet
    Dim wsLookup As Worksheet
    Dim wsOutput As Worksheet
    Dim dict As Object
    Dim lastRowData As Long, lastRowLookup As Long
    Dim lastColData As Long
    Dim i As Long, j As Long
    Dim caseID As Variant
    Dim outputRow As Long

    ' Set sheets
    Set wsData = ThisWorkbook.Sheets("PerCaseBreakdown")
    Set wsLookup = ThisWorkbook.Sheets("Casestocheck")

    ' Create or set Output sheet
    On Error Resume Next
    Set wsOutput = ThisWorkbook.Sheets("Output")
    On Error GoTo 0

    If wsOutput Is Nothing Then
        Set wsOutput = ThisWorkbook.Sheets.Add
        wsOutput.name = "Output"
    Else
        wsOutput.Cells.Clear
    End If

    ' Get dimensions
    lastRowData = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    lastRowLookup = wsLookup.Cells(wsLookup.Rows.Count, 3).End(xlUp).Row ' Case ID in Column C
    lastColData = wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column

    ' Create dictionary
    Set dict = CreateObject("Scripting.Dictionary")

    ' Load lookup values from Casestocheck (columns E,F,I,J,K,L,M)
    For i = 2 To lastRowLookup
        caseID = wsLookup.Cells(i, 3).Value ' Column C = Case ID
        
        If Not dict.Exists(caseID) Then
            dict.Add caseID, Array( _
                wsLookup.Cells(i, 5).Value, _   ' Column E
                wsLookup.Cells(i, 6).Value, _   ' Column F
                wsLookup.Cells(i, 9).Value, _   ' Column I
                wsLookup.Cells(i, 10).Value, _  ' Column J
                wsLookup.Cells(i, 11).Value, _  ' Column K
                wsLookup.Cells(i, 12).Value, _  ' Column L
                wsLookup.Cells(i, 13).Value _   ' Column M
            )
        End If
    Next i

    ' Copy all headers from PerCaseBreakdown
    ' Headers: CaseID, EmailIn_InPeriod, EmailIn_OutPeriod, EmailOut_InPeriod, 
    '          EmailOut_OutPeriod, TotalEmailIn, TotalEmailOut, PhoneIn_In, 
    '          PhoneIn_Out, PhoneOut_In, PhoneOut_Out, Chat_In, Chat_Out, 
    '          WhatsApp_In, WhatsApp_Out, Total_In, Total_Out, TotalInteractions, PeriodName
    wsData.Rows(1).Copy wsOutput.Rows(1)

    ' Add headers from lookup sheet (E,F,I,J,K,L,M)
    wsOutput.Cells(1, lastColData + 1).Value = wsLookup.Cells(1, 5).Value
    wsOutput.Cells(1, lastColData + 2).Value = wsLookup.Cells(1, 6).Value
    wsOutput.Cells(1, lastColData + 3).Value = wsLookup.Cells(1, 9).Value
    wsOutput.Cells(1, lastColData + 4).Value = wsLookup.Cells(1, 10).Value
    wsOutput.Cells(1, lastColData + 5).Value = wsLookup.Cells(1, 11).Value
    wsOutput.Cells(1, lastColData + 6).Value = wsLookup.Cells(1, 12).Value
    wsOutput.Cells(1, lastColData + 7).Value = wsLookup.Cells(1, 13).Value

    outputRow = 2

    ' Merge process - copy all data rows
    For i = 2 To lastRowData

        ' Copy entire row from PerCaseBreakdown
        wsOutput.Cells(outputRow, 1).Resize(1, lastColData).Value = _
            wsData.Cells(i, 1).Resize(1, lastColData).Value

        ' Get Case ID from first column
        caseID = wsData.Cells(i, 1).Value

        ' Append 7 lookup columns
        If dict.Exists(caseID) Then
            wsOutput.Cells(outputRow, lastColData + 1).Resize(1, 7).Value = dict(caseID)
        Else
            ' Fill with "No Match" if Case ID not found in lookup
            For j = 1 To 7
                wsOutput.Cells(outputRow, lastColData + j).Value = "No Match"
            Next j
        End If

        outputRow = outputRow + 1

    Next i

    ' Format output sheet
    With wsOutput.UsedRange
        .Font.Name = "Calibri"
        .Font.Size = 11
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' AutoFit columns
    wsOutput.Columns("A:" & Chr(64 + lastColData + 7)).AutoFit

    ' Freeze header row
    wsOutput.Range("A2").Select
    ActiveWindow.FreezePanes = True

    MsgBox "Merge complete! " & (outputRow - 2) & " records processed.", vbInformation, "Merge Status"

End Sub
