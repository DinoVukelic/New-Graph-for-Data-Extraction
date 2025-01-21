Sub ProcessAllSheets()
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
                
                ' Strip numbers from the machine name
                cleanMachineName = RemoveNumbers(machineName)
                
                ' Check if the machine name matches any category
                For Each category In machineCategories.Keys
                    If InStr(1, cleanMachineName, machineCategories(category), vbTextCompare) > 0 Then
                        ' Validate and add the time if it's in a valid format
                        If IsNumeric(TimeValue("00:" & timeValue)) Then
                            On Error Resume Next
                            machineTimes(category).Add CDate("00:" & timeValue)
                            On Error GoTo 0
                        End If
                        Exit For
                    End If
                Next category
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

' Function to remove numbers from machine names
Function RemoveNumbers(ByVal inputString As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.Pattern = "[0-9]"
    RemoveNumbers = regex.Replace(inputString, "")
End Function
