Option Explicit

Sub ProcessAllSheetsExcludeHiddenRowsAndColumns()
    Dim ws As Worksheet
    Dim reportWs As Worksheet
    
    Dim machineTimes As Object      ' Dictionary: category => Collection (Double times)
    Dim machineCategories As Object ' Dictionary: "VIRTXEN"/"VDIST"/"VIRTPPC"
    
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
    machineCategories.Add "VIRTXEN", "VIRTXEN"
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
            
            Debug.Print "Processing sheet: " & ws.Name
            
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
                            
                            ' Reset numericTime for this column
                            numericTime = 0
                            
                            ' Get time value from row 14
                            On Error Resume Next
                            If Not IsEmpty(ws.Cells(14, i)) Then
                                cellValue = ws.Cells(14, i).Value
                            Else
                                cellValue = vbNullString
                            End If
                            On Error GoTo 0
                            
                            ' Debug what we are about to parse
                            Debug.Print "  Column " & i & ": machine=" & machineName & _
                                        ", cleanMachine=" & cleanMachineName & _
                                        ", row14=[" & CStr(cellValue) & "]"
                            
                            ' Safely parse the time (returns 0 if invalid)
                            numericTime = SafeParseTime(cellValue)
                            
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
    
    ' Pattern allows "VIRTXEN", "VDIST", "VIRTPPC" 
    ' plus any trailing letters, digits, or underscores.
    regex.Pattern = "^(VIRTXEN\w*|VDIST\w*|VIRTPPC\w*)"
    regex.Global = False
    regex.IgnoreCase = False
    
    If regex.Test(inputString) Then
        ' Return entire match (e.g., "VIRTXEN6" or "VDIST305")
        ExtractMachineName = regex.Execute(inputString)(0)
    Else
        ExtractMachineName = ""
    End If
End Function

'-----------------------------------
' Safe function to parse time values
'-----------------------------------
Function SafeParseTime(ByVal textValue As Variant) As Double
    Dim tmpDate As Date
    
    On Error GoTo FailSafe
    
    ' If it's already a date/time, just convert
    If IsDate(textValue) Then
        SafeParseTime = CDbl(textValue)
        Exit Function
    End If
    
    ' If it's a 6-digit numeric string like "002508", parse manually
    If VarType(textValue) = vbString Then
        If Len(textValue) = 6 And IsNumeric(textValue) Then
            SafeParseTime = TimeSerial(Left(textValue, 2), Mid(textValue, 3, 2), Right(textValue, 2))
            Exit Function
        ElseIf InStr(textValue, ":") > 0 Then
            ' If it has a colon, attempt TimeValue
            tmpDate = TimeValue(textValue)
            SafeParseTime = CDbl(tmpDate)
            Exit Function
        End If
    End If

FailSafe:
    ' If parsing fails, or if we got here, return 0
    SafeParseTime = 0
End Function
