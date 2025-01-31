Option Explicit

Sub ProcessAllSheetsExcludeHiddenRowsAndColumns()

    Dim sh As Object         ' Generic sheet object (could be chart, macro sheet, etc.)
    Dim ws As Worksheet      ' Actual worksheet reference
    
    Dim reportWs As Worksheet
    
    Dim machineTimes As Object      ' Dictionary: category => Collection (Double times)
    Dim machineCategories As Object ' Dictionary: "VIRTXEN"/"VDIST"/"VIRTPPC"/"372WTB"
    
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
    machineCategories.Add "372WTB", "372WTB"
    
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
        .Range("D1").Value = "Sheet Name (Day_Month)"
        .Range("A1:D1").Font.Bold = True
    End With
    reportRow = 2
    
    '---------------------------------------------
    ' 3) Loop over ALL sheets, but only process Worksheet objects
    '---------------------------------------------
    For Each sh In ThisWorkbook.Sheets
        
        ' Only proceed if this is a normal Worksheet
        If TypeName(sh) = "Worksheet" Then
            Set ws = sh
            
            ' We check that the name matches "*_BBM_Export_Timings"
            ' AND also is NOT one of these excluded names
            If (ws.Name Like "*_BBM_Export_Timings") _
                And (ws.Name <> "DD_MM_BBM_Export_Timings") _
                And (ws.Name <> "DD_MM_BBS_Export_Timings") Then
                
                Debug.Print vbNewLine & ">>> Processing sheet: " & ws.Name
                
                ' Create a new dictionary for times
                Set machineTimes = CreateObject("Scripting.Dictionary")
                For Each category In machineCategories.Keys
                    Set timeCollection = New Collection
                    machineTimes.Add category, timeCollection
                Next category
                
                '---------------------------------------------
                ' 4) Process visible columns H to BQ (8 to 69)
                '---------------------------------------------
                For i = 8 To 69
                    ' Skip hidden columns
                    If Not ws.Columns(i).Hidden Then
                        
                        ' Safely read the machine name from row 1
                        If Not IsError(ws.Cells(1, i).Value) Then
                            cellValue = ws.Cells(1, i).Value
                        Else
                            cellValue = vbNullString
                        End If
                        
                        machineName = ""
                        If Not IsEmpty(cellValue) Then
                            machineName = Trim(CStr(cellValue))
                        End If
                        
                        ' Only process if row 1 has text
                        If Len(machineName) > 0 Then
                            ' Extract prefix: VIRTXEN, VDIST, VIRTPPC, or 372WTB
                            cleanMachineName = ExtractMachineName(machineName)
                            
                            Debug.Print "  Column " & i & _
                                        " => row1='" & machineName & _
                                        "' => Extracted='" & cleanMachineName & "'"
                            
                            If Len(cleanMachineName) > 0 Then
                                
                                ' Get time value from row 14 safely
                                If Not IsError(ws.Cells(14, i).Value) Then
                                    cellValue = ws.Cells(14, i).Value
                                Else
                                    cellValue = vbNullString
                                End If
                                
                                Debug.Print "    row14=[" & CStr(cellValue) & "]"
                                
                                ' Convert to numeric time
                                numericTime = SafeParseTime(cellValue)
                                
                                If numericTime > 0 Then
                                    Debug.Print "    numericTime recognized: " & numericTime
                                    
                                    ' Add to the correct category
                                    Select Case cleanMachineName
                                        Case "VIRTXEN"
                                            machineTimes("VIRTXEN").Add numericTime
                                            Debug.Print "    ==> Added to category: VIRTXEN"
                                        Case "VDIST"
                                            machineTimes("VDIST").Add numericTime
                                            Debug.Print "    ==> Added to category: VDIST"
                                        Case "VIRTPPC"
                                            machineTimes("VIRTPPC").Add numericTime
                                            Debug.Print "    ==> Added to category: VIRTPPC"
                                        Case "372WTB"
                                            machineTimes("372WTB").Add numericTime
                                            Debug.Print "    ==> Added to category: 372WTB"
                                    End Select
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
                Dim idx As Long
                Dim minTotalSec As Long, maxTotalSec As Long
                Dim minMinutes As Long, minSeconds As Long
                Dim maxMinutes As Long, maxSeconds As Long
                Dim categoryNameForReport As String
                
                For Each category In machineCategories.Keys
                    If machineTimes(category).Count > 0 Then
                        minTime = 999999    ' Initialize to a very high number
                        maxTime = 0
                        
                        ' Find min and max times
                        For idx = 1 To machineTimes(category).Count
                            ' Each item should be a Double
                            numericTime = CDbl(machineTimes(category).Item(idx))
                            
                            If numericTime < minTime Then minTime = numericTime
                            If numericTime > maxTime Then maxTime = numericTime
                        Next idx
                        
                        ' Convert fraction-of-day to total minutes/seconds
                        minTotalSec = CLng(minTime * 86400)
                        maxTotalSec = CLng(maxTime * 86400)
                        
                        minMinutes = minTotalSec \ 60
                        minSeconds = minTotalSec Mod 60
                        maxMinutes = maxTotalSec \ 60
                        maxSeconds = maxTotalSec Mod 60
                        
                        ' Label for the report
                        If category = "VIRTPPC" Then
                            categoryNameForReport = "Citrix - VIRTPPC"
                        Else
                            categoryNameForReport = category
                        End If
                        
                        ' Write to the report sheet
                        With reportWs
                            .Cells(reportRow, 1).Value = categoryNameForReport
                            .Cells(reportRow, 2).Value = CStr(minMinutes) & ":" & Format(minSeconds, "00")
                            .Cells(reportRow, 3).Value = CStr(maxMinutes) & ":" & Format(maxSeconds, "00")
                            .Cells(reportRow, 4).Value = ws.Name
                        End With
                        
                        Debug.Print "==> " & ws.Name & ", Category=" & category & _
                                    ", Min=" & minMinutes & ":" & Format(minSeconds, "00") & _
                                    ", Max=" & maxMinutes & ":" & Format(maxSeconds, "00")
                        
                        reportRow = reportRow + 1
                    Else
                        Debug.Print "==> " & ws.Name & ", Category=" & category & " => no times found."
                    End If
                Next category
                
            End If ' end of \"If ws.Name Like...\"
        End If ' end of \"If TypeName(sh) = 'Worksheet'...\"
        
    Next sh
    
    '---------------------------------------------
    ' 6) Format the report (only if we have rows)
    '---------------------------------------------
    If reportWs Is Nothing Then
        MsgBox "Error: Report worksheet not found"
        GoTo Cleanup
    End If
    
    Dim lastRow As Long
    lastRow = reportRow - 1
    
    If lastRow >= 2 Then
        ' AutoFit columns
        reportWs.Columns("A:D").AutoFit
        
        ' Add borders
        With reportWs.Range("A1:D" & lastRow)
            .Borders.LineStyle = xlContinuous
        End With
        
        ' Set alignment
        reportWs.Range("B:C").HorizontalAlignment = xlHAlignLeft
        reportWs.Range("D:D").HorizontalAlignment = xlHAlignCenter
    End If
    
Cleanup:
    ' Restore application settings
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "Processing complete. Check the 'Machine Times Report' sheet and VBA Immediate Window for details."

End Sub

'================================================
' Function: ExtractMachineName
'   - Returns "VIRTXEN", "VDIST", "VIRTPPC", or "372WTB"
'   - Ignores trailing characters
'================================================
Function ExtractMachineName(ByVal inputString As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    regex.Pattern = "^(VIRTXEN|VDIST|VIRTPPC|372WTB)"
    regex.IgnoreCase = False
    regex.Global = False
    
    If regex.Test(inputString) Then
        ExtractMachineName = regex.Execute(inputString)(0)
    Else
        ExtractMachineName = ""
    End If
End Function

'-----------------------------------
' Safe function to parse time values
' Returns 0 if not parseable.
'-----------------------------------
Function SafeParseTime(ByVal textValue As Variant) As Double
    Dim tmpDate As Date
    
    ' 1) If cell is an error or is Null, skip
    If IsError(textValue) Or IsNull(textValue) Then
        SafeParseTime = 0
        Exit Function
    End If
    
    ' 2) Directly a Date/Time value
    If IsDate(textValue) Then
        SafeParseTime = CDbl(textValue)
        Exit Function
    End If
    
    ' 3) If numeric fraction between 0 and 1, treat as fraction of a day
    If IsNumeric(textValue) Then
        If textValue >= 0 And textValue < 1 Then
            SafeParseTime = CDbl(textValue)
            Exit Function
        End If
    End If
    
    ' 4) If it's a 6-digit numeric string \"HHMMSS\", parse via TimeSerial
    If VarType(textValue) = vbString Then
        Dim s As String
        s = Trim(textValue)
        
        If Len(s) = 6 And IsNumeric(s) Then
            On Error GoTo ParseFail
            SafeParseTime = TimeSerial(Left(s, 2), Mid(s, 3, 2), Right(s, 2))
            Exit Function
        ElseIf InStr(s, \":\") > 0 Then
            ' If it has a colon, interpret as \"hh:mm\" or \"hh:mm:ss\"
            On Error GoTo ParseFail
            tmpDate = TimeValue(s)
            SafeParseTime = CDbl(tmpDate)
            Exit Function
        End If
    End If
    
ParseFail:
    ' If we can't parse, return 0
    SafeParseTime = 0
End Function
