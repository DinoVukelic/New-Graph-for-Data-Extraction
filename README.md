Sub ProcessAllSheetsDebug()
    Dim ws As Worksheet
    Dim reportWs As Worksheet
    Dim machineTimes As Object
    Dim machineCategories As Object
    Dim colStart As Long, colEnd As Long
    Dim i As Long
    Dim machineName As String
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
    
    ' Create a new sheet for the report
    On Error Resume Next
    Set reportWs = ThisWorkbook.Sheets("Machine Times Report")
    If reportWs Is Nothing Then
        Set reportWs = ThisWorkbook.Sheets.Add
        reportWs.Name = "Machine Times Report"
    End If
    On Error GoTo 0
    reportWs.Cells.Clear
    reportWs.Range("A1:D1").Value = Array("Machine Category", "Min Time (mm:ss)", "Max Time (mm:ss)", "Sheet Name")
    reportRow = 2
    
    ' Process each sheet matching the pattern
    For Each ws In ThisWorkbook.Sheets
        If ws.Name Like "*_BBM_Export_Timings" Then
            ' Get the start and end columns for the data
            colStart = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
            colEnd = ws.Cells(1, 1).End(xlToRight).Column
            
            ' Initialize dictionary to store times for each category
            Set machineTimes = CreateObject("Scripting.Dictionary")
            For Each category In machineCategories.Keys
                machineTimes.Add category, New Collection
            Next category
            
            ' Loop through the columns
            For i = colStart To colEnd
                machineName = ws.Cells(1, i).Value
                timeValue = ws.Cells(14, i).Value
                
                ' Debug: Log column and cell values
                Debug.Print "Processing column: " & i & ", Machine Name: " & machineName & ", Time: " & timeValue
                
                ' Ignore empty machine names and times
                If IsEmpty(machineName) Or IsEmpty(timeValue) Then
                    Debug.Print "Skipping column " & i & " due to empty value."
                    GoTo SkipColumn
                End If
                
                ' Extract the base machine name
                cleanMachineName = ExtractMachineName(machineName)
                Debug.Print "Extracted Machine Name: " & cleanMachineName
                
                ' Check if the machine name matches any category
                For Each category In machineCategories.Keys
                    If cleanMachineName = machineCategories(category) Then
                        ' Validate and process the time
                        If IsDate(timeValue) Then
                            On Error Resume Next
                            machineTimes(category).Add TimeValueToMinutesSeconds(timeValue)
                            On Error GoTo 0
                        Else
                            Debug.Print "Invalid time format in column " & i & ": " & timeValue
                        End If
                        Exit For
                    End If
                Next category
SkipColumn:
            Next i
            
            ' Calculate min and max times for each category
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
                    reportWs.Cells(reportRow, 2).Value = Format(minTime, "mm:ss")
                    reportWs.Cells(reportRow, 3).Value = Format(maxTime, "mm:ss")
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

' Function to extract the base machine name (e.g., "VIRTIXEN" from "VIRTIXEN630S")
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

' Function to convert time values to mm:ss format
Function TimeValueToMinutesSeconds(ByVal timeValue As Variant) As Date
    Dim totalSeconds As Double
    Dim minutes As Long
    Dim seconds As Long

    ' Calculate total seconds
    totalSeconds = timeValue * 86400 ' Convert Excel time to total seconds

    ' Extract minutes and seconds
    minutes = Int(totalSeconds / 60)
    seconds = totalSeconds Mod 60

    ' Return as mm:ss
    TimeValueToMinutesSeconds = TimeSerial(0, minutes, seconds)
End Function
