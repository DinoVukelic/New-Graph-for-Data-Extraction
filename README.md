Option Explicit

Sub ProcessAllSheetsExcludeHiddenRowsAndColumns()
    Dim ws As Worksheet
    Dim reportWs As Worksheet
    
    Dim machineTimes As Object      ' Dictionary: category => Collection (Double times)
    Dim machineCategories As Object ' Dictionary: "VIRTIXEN"/"VDIST"/"VIRTPPC"
    
    Dim i As Long
    Dim machineName As Variant
    Dim cleanMachineName As String
    
    Dim timeValue As Variant
    Dim numericTime As Double
    
    Dim category As Variant
    Dim minTime As Double
    Dim maxTime As Double
    
    Dim reportRow As Long
    Dim timeArray() As Double
    Dim timeIndex As Long
    
    '---------------------------
    ' 1) Initialize categories
    '---------------------------
    Set machineCategories = CreateObject("Scripting.Dictionary")
    machineCategories.Add "VIRTIXEN", "VIRTIXEN"
    machineCategories.Add "VDIST", "VDIST"
    machineCategories.Add "VIRTPPC", "VIRTPPC"
    
    '---------------------------------------------
    ' 2) Create/clear the "Machine Times Report"
    '---------------------------------------------
    On Error Resume Next
    Set reportWs = ThisWorkbook.Sheets("Machine Times Report")
    On Error GoTo 0
    
    If reportWs Is Nothing Then
        Set reportWs = ThisWorkbook.Sheets.Add
        reportWs.Name = "Machine Times Report"
    Else
        reportWs.Cells.Clear
    End If
    
    ' Add headers
    With reportWs
        .Range("A1").Value = "Machine Category"
        .Range("B1").Value = "Min Time (hh:mm:ss)"
        .Range("C1").Value = "Max Time (hh:mm:ss)"
        .Range("D1").Value = "Sheet Name"
    End With
    reportRow = 2
    
    '---------------------------------------------
    ' 3) Loop each sheet named "*_BBM_Export_Timings"
    '---------------------------------------------
    For Each ws In ThisWorkbook.Sheets
        If ws.Name Like "*_BBM_Export_Timings" Then
            
            ' Create a new dictionary for times
            Set machineTimes = CreateObject("Scripting.Dictionary")
            For Each category In machineCategories.Keys
                machineTimes.Add category, New Collection
            Next category
            
            '---------------------------------------------
            ' 4) Process columns H to BQ
            '---------------------------------------------
            For i = 8 To 69  ' 8=H, 69=BQ
                ' Use CStr to safely handle cell values
                If Not IsEmpty(ws.Cells(7, i)) Then
                    machineName = CStr(ws.Cells(7, i).Text)
                    
                    If Not IsEmpty(ws.Cells(14, i)) Then
                        timeValue = ws.Cells(14, i).Text
                        
                        ' Skip if machineName is empty
                        If Len(machineName) > 0 Then
                            ' Extract the base name
                            cleanMachineName = ExtractMachineName(machineName)
                            
                            If Len(cleanMachineName) > 0 Then
                                ' Handle time conversion with error checking
                                On Error Resume Next
                                If IsDate(timeValue) Then
                                    numericTime = CDbl(TimeValue(CStr(timeValue)))
                                Else
                                    numericTime = CDbl(TimeValue(CStr(timeValue)))
                                End If
                                On Error GoTo 0
                                
                                ' Only add valid times
                                If numericTime > 0 Then
                                    For Each category In machineCategories.Keys
                                        If cleanMachineName = machineCategories(category) Then
                                            machineTimes(category).Add numericTime
                                            Exit For
                                        End If
                                    Next category
                                End If
                            End If
                        End If
                    End If
                End If
            Next i
            
            '---------------------------------------------
            ' 5) Calculate and write min/max times
            '---------------------------------------------
            For Each category In machineCategories.Keys
                If machineTimes(category).Count > 0 Then
                    ReDim timeArray(1 To machineTimes(category).Count)
                    
                    ' Safely copy times to array
                    For timeIndex = 1 To machineTimes(category).Count
                        timeArray(timeIndex) = CDbl(machineTimes(category).Item(timeIndex))
                    Next timeIndex
                    
                    ' Calculate min/max
                    minTime = Application.WorksheetFunction.Min(timeArray)
                    maxTime = Application.WorksheetFunction.Max(timeArray)
                    
                    ' Write to report with error handling
                    With reportWs
                        .Cells(reportRow, 1).Value = CStr(category)
                        .Cells(reportRow, 2).Value = Format(minTime, "hh:mm:ss")
                        .Cells(reportRow, 3).Value = Format(maxTime, "hh:mm:ss")
                        .Cells(reportRow, 4).Value = ws.Name
                    End With
                    reportRow = reportRow + 1
                End If
            Next category
        End If
    Next ws
    
    '---------------------------------------------
    ' 6) Format the report
    '---------------------------------------------
    reportWs.Columns.AutoFit
    MsgBox "Processing complete. Check the 'Machine Times Report' sheet."
End Sub

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
