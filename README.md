Option Explicit

Sub Machine_Times_Report()

    Dim sh As Object         ' Generic sheet object (could be chart, macro sheet, etc.)
    Dim ws As Worksheet      ' Actual worksheet reference
    Dim reportWs As Worksheet
    
    Dim machineTimes As Object  ' Dictionary: full machine name => Collection (Double times)
    Dim machineCat As Object    ' Dictionary: full machine name => Machine Category
    
    Dim i As Long
    Dim machineNameFull As String
    Dim machinePrefix As String
    Dim cellValue As Variant
    Dim numericTime As Double
    Dim category As String
    
    Dim reportRow As Long
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    '---------------------------------------------
    ' 1) Create/clear the "Machine Times Report" sheet
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
    
    ' Add headers with the new column order:
    ' A = Machine Name, B = Machine Category, C = Min Time, D = Max Time, E = Sheet Name
    With reportWs
        .Range("A1").Value = "Machine Name"
        .Range("B1").Value = "Machine Category"
        .Range("C1").Value = "Min Time (mm:ss)"
        .Range("D1").Value = "Max Time (mm:ss)"
        .Range("E1").Value = "Sheet Name (Day_Month)"
        .Range("A1:E1").Font.Bold = True
    End With
    reportRow = 2
    
    '---------------------------------------------
    ' 2) Loop over ALL sheets, but only process Worksheet objects
    '---------------------------------------------
    For Each sh In ThisWorkbook.Sheets
        
        ' Only proceed if this is a normal Worksheet
        If TypeName(sh) = "Worksheet" Then
            Set ws = sh
            
            ' Process only sheets with names matching "*_BBM_Export_Timings"
            ' and not the excluded names
            If (ws.Name Like "*_BBM_Export_Timings") And _
               (ws.Name <> "DD_MM_BBM_Export_Timings") And _
               (ws.Name <> "DD_MM_BBS_Export_Timings") Then
                
                Debug.Print vbNewLine & ">>> Processing sheet: " & ws.Name
                
                ' Create new dictionaries for each sheet.
                ' machineTimes: key = full machine name, value = Collection of numeric times
                ' machineCat: key = full machine name, value = Machine Category
                Set machineTimes = CreateObject("Scripting.Dictionary")
                Set machineCat = CreateObject("Scripting.Dictionary")
                
                '---------------------------------------------
                ' 3) Process visible columns H to BQ (columns 8 to 69)
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
                        
                        machineNameFull = ""
                        If Not IsEmpty(cellValue) Then
                            machineNameFull = Trim(CStr(cellValue))
                        End If
                        
                        ' Only process if row 1 has text
                        If Len(machineNameFull) > 0 Then
                            
                            ' Get the prefix (VIRTXEN, VDIST, VIRTPPC, or 372WTB) from the machine name
                            machinePrefix = ExtractMachineName(machineNameFull)
                            
                            ' Determine Machine Category based on the prefix
                            Select Case machinePrefix
                                Case "VIRTPPC", "VIRTXEN"
                                    category = "Citrix"
                                Case "VDIST"
                                    category = "Virtual Machine"
                                Case "372WTB"
                                    category = "Physical Machine"
                                Case Else
                                    category = "Other"
                            End Select
                            
                            Debug.Print "  Column " & i & " => row1='" & machineNameFull & "'; Prefix='" & machinePrefix & "'; Category='" & category & "'"
                            
                            ' Initialize collection and store category if this machine name has not been processed yet
                            If Not machineTimes.Exists(machineNameFull) Then
                                machineTimes.Add machineNameFull, New Collection
                                machineCat.Add machineNameFull, category
                            End If
                            
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
                                machineTimes(machineNameFull).Add numericTime
                            Else
                                Debug.Print "    --> Time not parsed or invalid."
                            End If
                        Else
                            Debug.Print "  Column " & i & " => row1 is empty or error; skipped."
                        End If
                    End If
                Next i
                
                '---------------------------------------------
                ' 4) Calculate and write min/max times per machine (only if times were recorded)
                '---------------------------------------------
                Dim key As Variant
                Dim idx As Long
                Dim minTime As Double, maxTime As Double
                Dim minTotalSec As Long, maxTotalSec As Long
                Dim minMinutes As Long, minSeconds As Long
                Dim maxMinutes As Long, maxSeconds As Long
                
                For Each key In machineTimes.Keys
                    If machineTimes(key).Count > 0 Then
                        minTime = 999999    ' Initialize to a very high number
                        maxTime = 0
                        
                        ' Loop through all times for this machine to get the min and max
                        For idx = 1 To machineTimes(key).Count
                            numericTime = CDbl(machineTimes(key).Item(idx))
                            
                            If numericTime < minTime Then minTime = numericTime
                            If numericTime > maxTime Then maxTime = numericTime
                        Next idx
                        
                        ' Convert fraction-of-day to total seconds, then calculate minutes and seconds
                        minTotalSec = CLng(minTime * 86400)
                        maxTotalSec = CLng(maxTime * 86400)
                        
                        minMinutes = minTotalSec \ 60
                        minSeconds = minTotalSec Mod 60
                        maxMinutes = maxTotalSec \ 60
                        maxSeconds = maxTotalSec Mod 60
                        
                        ' Write results to the report sheet:
                        ' Column A: Machine Name (full)
                        ' Column B: Machine Category (as determined)
                        ' Column C: Min Time (mm:ss)
                        ' Column D: Max Time (mm:ss)
                        ' Column E: Sheet Name
                        With reportWs
                            .Cells(reportRow, 1).Value = key
                            .Cells(reportRow, 2).Value = machineCat(key)
                            .Cells(reportRow, 3).Value = CStr(minMinutes) & ":" & Format(minSeconds, "00")
                            .Cells(reportRow, 4).Value = CStr(maxMinutes) & ":" & Format(maxSeconds, "00")
                            .Cells(reportRow, 5).Value = ws.Name
                        End With
                        
                        Debug.Print "==> " & ws.Name & ", Machine='" & key & "', Category='" & machineCat(key) & _
                                    "', Min=" & minMinutes & ":" & Format(minSeconds, "00") & _
                                    ", Max=" & maxMinutes & ":" & Format(maxSeconds, "00")
                        
                        reportRow = reportRow + 1
                    Else
                        Debug.Print "==> " & ws.Name & ", Machine='" & key & "' => no valid times found."
                    End If
                Next key
                
            End If ' End If sheet name matches
        End If ' End If TypeName(ws) = "Worksheet"
        
    Next sh
    
    '---------------------------------------------
    ' 5) Format the report (only if we have rows)
    '---------------------------------------------
    If reportWs Is Nothing Then
        MsgBox "Error: Report worksheet not found"
        GoTo Cleanup
    End If
    
    Dim lastRow As Long
    lastRow = reportRow - 1
    
    If lastRow >= 2 Then
        ' AutoFit columns
        reportWs.Columns("A:E").AutoFit
        
        ' Add borders
        With reportWs.Range("A1:E" & lastRow)
            .Borders.LineStyle = xlContinuous
        End With
        
        ' Set alignment
        reportWs.Range("C:D").HorizontalAlignment = xlHAlignLeft
        reportWs.Range("E:E").HorizontalAlignment = xlHAlignCenter
        
        ' Apply AutoFilter to the header row
        reportWs.Range("A1:E1").AutoFilter
    End If
    
Cleanup:
    ' Restore application settings
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "Processing complete. Check the 'Machine Times Report' sheet and the VBA Immediate Window for details."

End Sub

'================================================
' Function: ExtractMachineName
'   - Returns "VIRTXEN", "VDIST", "VIRTPPC", or "372WTB" if found at the start.
'   - If not, returns an empty string.
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
    
    ' 4) If it's a 6-digit numeric string "HHMMSS", parse via TimeSerial
    If VarType(textValue) = vbString Then
        Dim s As String
        s = Trim(textValue)
        
        If Len(s) = 6 And IsNumeric(s) Then
            On Error GoTo ParseFail
            SafeParseTime = TimeSerial(Left(s, 2), Mid(s, 3, 2), Right(s, 2))
            Exit Function
        ElseIf InStr(s, ":") > 0 Then
            ' If it has a colon, interpret as "hh:mm" or "hh:mm:ss"
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
