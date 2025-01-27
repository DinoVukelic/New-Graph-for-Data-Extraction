Option Explicit

Sub ProcessAllSheetsExcludeHiddenRowsAndColumns()
    Dim ws As Worksheet
    Dim reportWs As Worksheet
    
    Dim machineTimes As Object      ' Dictionary: category => Collection (Double times)
    Dim machineCategories As Object ' Dictionary: "VIRTIXEN"/"VDIST"/"VIRTPPC"
    
    Dim i As Long
    Dim machineName As String
    Dim cleanMachineName As String
    
    Dim cellValue As Variant
    Dim numericTime As Double
    
    Dim category As Variant
    Dim minTime As Double
    Dim maxTime As Double
    
    Dim reportRow As Long
    Dim timeCollection As Collection
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
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
            ' 4) Process visible columns H to BQ
            '---------------------------------------------
            For i = 8 To 69  ' 8=H, 69=BQ
                ' Skip hidden columns
                If Not ws.Columns(i).Hidden Then
                    ' Get machine name from row 1
                    machineName = ""
                    On Error Resume Next
                    If Not IsEmpty(ws.Cells(1, i)) Then
                        cellValue = ws.Cells(1, i).Value
                        If Not IsError(cellValue) Then
                            machineName = Trim(CStr(cellValue))
                        End If
                    End If
                    On Error GoTo 0
                    
                    ' Process if we have a valid machine name
                    If Len(machineName) > 0 Then
                        cleanMachineName = ExtractMachineName(machineName)
                        
                        If Len(cleanMachineName) > 0 Then
                            ' Get time value from row 14
                            On Error Resume Next
                            If Not IsEmpty(ws.Cells(14, i)) Then
                                cellValue = ws.Cells(14, i).Value
                                
                                ' Handle different time formats
                                If Not IsError(cellValue) Then
                                    If IsDate(cellValue) Then
                                        numericTime = CDbl(cellValue)
                                    ElseIf VarType(cellValue) = vbString Then
                                        ' Try to parse time string
                                        If InStr(cellValue, ":") > 0 Then
                                            numericTime = TimeValue(cellValue)
                                        End If
                                    End If
                                    
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
                            On Error GoTo 0
                        End If
                    End If
                End If
            Next i
            
            '---------------------------------------------
            ' 5) Calculate and write min/max times
            '---------------------------------------------
            For Each category In machineCategories.Keys
                If machineTimes(category).Count > 0 Then
                    minTime = 1
                    maxTime = 0
                    
                    ' Find min and max times
                    For i = 1 To machineTimes(category).Count
                        On Error Resume Next
                        numericTime = machineTimes(category).Item(i)
                        If Err.Number = 0 Then
                            If numericTime < minTime Then minTime = numericTime
                            If numericTime > maxTime Then maxTime = numericTime
                        End If
                        On Error GoTo 0
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
        If reportRow > 2 Then
            .Range("A1:D" & reportRow - 1).Borders.LineStyle = xlContinuous
        End If
    End With
    
    Application.Calculation = xlCalculationAutomatic
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
