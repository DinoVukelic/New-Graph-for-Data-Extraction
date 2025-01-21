Sub ProcessAllSheetsExcludeHiddenRowsAndColumns()
    Dim ws As Worksheet
    Dim reportWs As Worksheet
    Dim machineTimes As Object
    Dim machineCategories As Object
    Dim colStart As Long, colEnd As Long
    Dim i As Long
    Dim machineName As Variant
    Dim cleanMachineName As String
    Dim timeValue As Variant
    Dim category As Variant
    Dim minTime As Date, maxTime As Date
    Dim reportRow As Long
    Dim timeArray() As Date
    Dim timeIndex As Long

    ' Initialize machine categories
    Set machineCategories = CreateObject("Scripting.Dictionary")
    machineCategories.Add "VIRTIXEN", "VIRTIXEN"
    machineCategories.Add "VDIST", "VDIST"
    machineCategories.Add "VIRTPPC", "VIRTPPC"

    ' Create or clear the "Machine Times Report" sheet
    On Error Resume Next
    Set reportWs = ThisWorkbook.Sheets("Machine Times Report")
    If reportWs Is Nothing Then
        Set reportWs = ThisWorkbook.Sheets.Add
        reportWs.Name = "Machine Times Report"
    End If
    On Error GoTo 0

    reportWs.Cells.Clear
    reportWs.Cells(1, 1).Value = "Machine Category"
    reportWs.Cells(1, 2).Value = "Min Time (hh:mm:ss)"
    reportWs.Cells(1, 3).Value = "Max Time (hh:mm:ss)"
    reportWs.Cells(1, 4).Value = "Sheet Name"
    reportRow = 2

    ' Process each sheet matching the pattern
    For Each ws In ThisWorkbook.Sheets
        Debug.Print "Processing sheet: " & ws.Name
        If ws.Name Like "*_BBM_Export_Timings" Then
            Debug.Print "Matched sheet: " & ws.Name

            ' Define start and end columns explicitly
            colStart = ws.Columns("H").Column ' Start at column H
            colEnd = ws.Columns("BQ").Column ' End at column BQ

            ' Initialize dictionary to store times for each category
            Set machineTimes = CreateObject("Scripting.Dictionary")
            For Each category In machineCategories.Keys
                machineTimes.Add category, New Collection
            Next category

            ' Loop through the columns
            For i = colStart To colEnd
                ' Read values from the sheet
                machineName = ws.Cells(1, i).Value
                timeValue = ws.Cells(14, i).Value

                ' Debug: Log the cell values
                Debug.Print "Processing column: " & i & ", Machine Name: " & machineName & ", Time: " & timeValue

                ' Skip columns if the machine name is error or empty
                If IsError(machineName) Or IsEmpty(machineName) Then
                    Debug.Print "Skipping column " & i & " due to empty/error machine name."
                    GoTo SkipColumn
                End If

                ' Extract the base machine name
                cleanMachineName = ExtractMachineName(CStr(machineName))
                If cleanMachineName = "" Then
                    Debug.Print "Skipping column " & i & " (invalid machine name): " & machineName
                    GoTo SkipColumn
                End If

                ' Skip columns if time value is error, empty, or not a date
                If IsError(timeValue) Or IsEmpty(timeValue) Or Not IsDate(timeValue) Then
                    Debug.Print "Skipping column " & i & " (invalid/empty time): " & timeValue
                    GoTo SkipColumn
                End If

                ' Add the time to the correct category if matched
                For Each category In machineCategories.Keys
                    If cleanMachineName = machineCategories(category) Then
                        On Error Resume Next
                        machineTimes(category).Add CDate(timeValue)
                        On Error GoTo 0
                        Exit For
                    End If
                Next category

SkipColumn:
            Next i

            ' Calculate min/max times for each category found on this sheet
            For Each category In machineCategories.Keys
                If machineTimes(category).Count > 0 Then
                    ' Convert the collection to an array
                    ReDim timeArray(1 To machineTimes(category).Count)
                    For timeIndex = 1 To machineTimes(category).Count
                        timeArray(timeIndex) = machineTimes(category).Item(timeIndex)
                    Next timeIndex

                    ' Calculate min and max times
                    minTime = Application.Min(timeArray)
                    maxTime = Application.Max(timeArray)

                    ' Write results to the report
                    reportWs.Cells(reportRow, 1).Value = category
                    reportWs.Cells(reportRow, 2).Value = Format(minTime, "hh:mm:ss")
                    reportWs.Cells(reportRow, 3).Value = Format(maxTime, "hh:mm:ss")
                    reportWs.Cells(reportRow, 4).Value = ws.Name
                    reportRow = reportRow + 1
                End If
            Next category
        End If
    Next ws

    ' Autofit columns in the report
    reportWs.Columns.AutoFit

    MsgBox "Processing complete. Check the 'Machine Times Report' sheet.", vbInformation
End Sub

' Extracts the base machine name (e.g., "VIRTIXEN" from "VIRTIXEN630S")
Function ExtractMachineName(ByVal inputString As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = False
    regex.Pattern = "^(VIRTIXEN|VDIST|VIRTPPC)"
    
    If regex.Test(inputString) Then
        ExtractMachineName = regex.Execute(inputString)(0)
    Else
        ExtractMachineName = ""
    End If
End Function
