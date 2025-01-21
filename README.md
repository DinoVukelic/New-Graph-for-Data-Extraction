Sub ProcessMachineTimes()
    Dim ws As Worksheet
    Dim machineTimes As Object
    Dim sheetPattern As String
    Dim machineCategories As Object
    Dim colStart As Long, colEnd As Long
    Dim i As Long
    Dim machineName As String
    Dim timeValue As Variant
    Dim category As Variant ' Ensure category is declared as Variant
    Dim minTime As Date, maxTime As Date
    Dim reportWs As Worksheet
    Dim reportRow As Long
    
    ' Initialize machine categories
    Set machineCategories = CreateObject("Scripting.Dictionary")
    machineCategories.Add "VIRTIXEN", "VIRTIXEN"
    machineCategories.Add "VDIST", "VDIST"
    machineCategories.Add "VIRTPPC", "VIRTPPC"
    
    ' Initialize sheet name pattern
    sheetPattern = "*_BBM_Export_Timings"
    
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
    
    ' Process each sheet
    For Each ws In ThisWorkbook.Sheets
        If ws.Name Like sheetPattern Then
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
                
                ' Check if machine name matches any category
                For Each category In machineCategories.Keys
                    If InStr(1, machineName, machineCategories(category), vbTextCompare) > 0 Then
                        ' Convert time to Excel time if valid
                        If IsDate("00:" & timeValue) Then
                            machineTimes(category).Add CDate("00:" & timeValue)
                        End If
                        Exit For
                    End If
                Next category
            Next i
            
            ' Calculate min and max times for each category
            For Each category In machineCategories.Keys
                If machineTimes(category).Count > 0 Then
                    minTime = Application.Min(machineTimes(category))
                    maxTime = Application.Max(machineTimes(category))
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
