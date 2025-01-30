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
        .Range("B1").Value = "Min Time (mm:ss)"
        .Range("C1").Value = "Max Time (mm:ss)"
        .Range("D1").Value = "Sheet Name"
        .Range("A1:D1").Font.Bold = True
    End With
    reportRow = 2
    
    '---------------------------------------------
    ' 3) Loop each sheet named "*_BBM_Export_Timings"
    '---------------------------------------------
    For Each ws In ThisWorkbook.Sheets
        If ws.Name Like "*_BBM_Export_Timings" Then
            
            Debug.Print vbNewLine & ">>> Processing sheet: " & ws.Name
            
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
                    
                    ' Only process if row 1 has text
                    If Len(machineName) > 0 Then
                        ' Extract prefix: VIRTXEN, VDIST, or VIRTPPC
                        cleanMachineName = ExtractMachineName(machineName)
                        
                        Debug.Print "  Column " & i & _
                                    " => row1='" & machineName & _
                                    "' => Extracted='" & cleanMachineName & "'"
                        
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
                            
                            Debug.Print "    row14=[" & CStr(cellValue) & "]"
                            
                            ' Convert to numeric time (if possible)
                            numericTime = SafeParseTime(cellValue)
                            
                            If numericTime > 0 Then
                                Debug.Print "    numericTime recognized: " & numericTime
                                
                                ' Add to the correct category
                                If cleanMachineName = "VIRTXEN" Then
                                    machineTimes("VIRTXEN").Add numericTime
                                    Debug.Print "    ==> Added to category: VIRTXEN"
                                ElseIf cleanMachineName = "VDIST" Then
                                    machineTimes("VDIST").Add numericTime
                                    Debug.Print "    ==> Added to category: VDIST"
                                ElseIf cleanMachineName = "VIRTPPC" Then
                                    machineTimes("VIRTPPC").Add numericTime
                                    Debug.Print "    ==> Added to category: VIRTPPC"
                                End If
                            Else
                                Debug.Print "    --> Time not parsed or invalid."
                            End If
                        Else
                            Debug.Print "    --> Machine name not recognized."
                        End If
                    Else
                        Debug.Print "  Column " & i & " => row1 is empty or error; skipped."
                    End If
                End If
            Next i
            
            '---------------------------------------------
            ' 5) Calculate and write min/max times in "mm:ss"
            '---------------------------------------------
            For Each category In machineCategories.Keys
                If machineTimes(category).Count > 0 Then
                    minTime = 1
                    maxTime = 0
                    
                    ' Find min and max times
                    Dim idx As Long
                    For idx = 1 To machineTimes(category).Count
                        On Error Resume Next
                        numericTime = machineTimes(category).Item(idx)
                        If Err.Number = 0 Then
                            If numericTime < minTime Then minTime = numericTime
                            If numericTime > maxTime Then maxTime = numericTime
                        End If
                        On Error GoTo 0
                    Next idx
                    
                    '-------------------------------------------------
                    ' Convert fraction-of-day to total minutes/seconds
                    '-------------------------------------------------
                    Dim minTotalSec As Long, maxTotalSec As Long
                    Dim minMinutes As Long, minSeconds As Long
                    Dim maxMinutes As Long, maxSeconds As Long
                    
                    minTotalSec = CLng(minTime * 86400) ' fraction of day -> total secs
                    maxTotalSec = CLng(maxTime * 86400)
                    
                    minMinutes = minTotalSec \ 60       ' integer division
                    minSeconds = minTotalSec Mod 60
                    maxMinutes = maxTotalSec \ 60
                    maxSeconds = maxTotalSec Mod 60
                    
                    ' Write "mm:ss" format with total minutes
                    With reportWs
                        .Cells(reportRow, 1).Value = CStr(category)
                        .Cells(reportRow, 2).Value = CStr(minMinutes) & ":" & Format(minSeconds, "00")
                        .Cells(reportRow, 3).Value = CStr(maxMinutes) & ":" & Format(maxSeconds, "00")
                        .Cells(reportRow, 4).Value = ws.Name
                    End With
                    
                    Debug.Print "==> " & ws.Name & ", Category=" & CStr(category) & _
                                ", Min=" & minMinutes & ":" & Format(minSeconds, "00") & _
                                ", Max=" & maxMinutes & ":" & Format(maxSeconds, "00")
                    
                    reportRow = reportRow + 1
                Else
                    Debug.Print "==> " & ws.Name & ", Category=" & category & " => no times found."
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
    MsgBox "Processing complete. Check the 'Machine Times Report' sheet and VBA Immediate Window for details."
End Sub

'================================================
' Function: ExtractMachineName
'   - Returns "VIRTXEN", "VDIST", or "VIRTPPC"
'   - Ignores trailing characters
'================================================
Function ExtractMachineName(ByVal inputString As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    regex.Pattern = "^(VIRTXEN|VDIST|VIRTPPC)"  ' match only exact prefix
    regex.IgnoreCase = False
    regex.Global = False
    
    If regex.Test(inputString) Then
        ' The entire match is e.g. "VIRTXEN"
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
    
    ' 1) If it's already a date/time
    If IsDate(textValue) Then
        SafeParseTime = CDbl(textValue)  ' convert to Double
        Exit Function
    End If
    
    ' 2) If numeric fraction between 0 and 1, treat as fraction of a 24-hour day
    If IsNumeric(textValue) Then
        If textValue >= 0 And textValue < 1 Then
            SafeParseTime = CDbl(textValue)
            Exit Function
        End If
    End If
    
    ' 3) If it's a 6-digit numeric string "002508", parse as hhmmss
    If VarType(textValue) = vbString Then
        If Len(textValue) = 6 And IsNumeric(textValue) Then
            SafeParseTime = TimeSerial(Left(textValue, 2), Mid(textValue, 3, 2), Right(textValue, 2))
            Exit Function
        ElseIf InStr(textValue, ":") > 0 Then
            ' If it has a colon, interpret as "hh:mm" or "hh:mm:ss"
            tmpDate = TimeValue(textValue)
            SafeParseTime = CDbl(tmpDate)
            Exit Function
        End If
    End If

FailSafe:
    ' If we can't parse, return 0
    SafeParseTime = 0
End Function
