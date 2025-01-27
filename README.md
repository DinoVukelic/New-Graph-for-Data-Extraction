Option Explicit

Sub ProcessAllSheetsExcludeHiddenRowsAndColumns()
    Dim ws As Worksheet
    Dim reportWs As Worksheet
    
    Dim machineTimes As Object      ' Dictionary: category => Collection (Double times)
    Dim machineCategories As Object ' Dictionary: "VIRTIXEN"/"VDIST"/"VIRTPPC"
    
    Dim i As Long
    Dim machineName As String
    Dim cleanMachineName As String
    
    Dim timeValue As String
    Dim numericTime As Double
    
    Dim category As Variant
    Dim minTime As Double
    Dim maxTime As Double
    
    Dim reportRow As Long
    Dim timeCollection As Collection
    
    Application.ScreenUpdating = False
    
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
        .Range("A1:D1").Font.Bold = True
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
                Set timeCollection = New Collection
                machineTimes.Add category, timeCollection
            Next category
            
            '---------------------------------------------
            ' 4) Process columns H to BQ
            '---------------------------------------------
            For i = 8 To 69  ' 8=H, 69=BQ
                On Error Resume Next
                machineName = ""
                If Not IsEmpty(ws.Cells(7, i)) Then
                    machineName = Trim(CStr(ws.Cells(7, i).Text))
                End If
                
                timeValue = ""
                If Not IsEmpty(ws.Cells(14, i)) Then
                    timeValue = Trim(CStr(ws.Cells(14, i).Text))
                End If
                On Error GoTo 0
                
                If Len(machineName) > 0 Then
                    cleanMachineName = ExtractMachineName(machineName)
                    
                    If Len(cleanMachineName) > 0 And Len(timeValue) > 0 Then
                        On Error Resume Next
                        ' Convert time value to Double
                        If IsDate(timeValue) Then
                            numericTime = CDbl(CDate(timeValue))
                        Else
                            numericTime = CDbl(CDate("1900-01-01 " & timeValue))
                        End If
                        
                        If Err.Number = 0 And numericTime > 0 Then
                            For Each category In machineCategories.Keys
                                If cleanMachineName = machineCategories(category) Then
                                    machineTimes(category).Add numericTime
                                    Exit For
                                End If
                            Next category
                        End If
                        On Error GoTo 0
                    End If
                End If
            Next i
            
            '---------------------------------------------
            ' 5) Calculate and write min/max times
            '---------------------------------------------
            For Each category In machineCategories.Keys
                If machineTimes(category).Count > 0 Then
                    minTime = 999999
                    maxTime = 0
                    
                    ' Find min and max times
                    For i = 1 To machineTimes(category).Count
                        numericTime = machineTimes(category).Item(i)
                        If numericTime < minTime Then minTime = numericTime
                        If numericTime > maxTime Then maxTime = numericTime
                    Next i
                    
                    ' Write to report
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
    With reportWs
        .Columns.AutoFit
        .Range("A1:D" & reportRow - 1).Borders.LineStyle = xlContinuous
    End With
    
    Application.ScreenUpdating = True
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
