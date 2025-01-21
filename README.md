Option Explicit

Sub ProcessAllSheetsExcludeHiddenRowsAndColumns()
    Dim ws As Worksheet
    Dim reportWs As Worksheet
    
    Dim machineTimes As Object      ' Dictionary: category => Collection
    Dim machineCategories As Object ' Dictionary: "VIRTIXEN" => "VIRTIXEN", etc.
    
    Dim colStart As Long, colEnd As Long
    Dim i As Long
    
    Dim machineName As Variant
    Dim cleanMachineName As String
    
    Dim timeValue As Variant
    Dim numericTime As Double       ' We'll store times as Double
    
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
    On Error GoTo CreateReportSheet
    Set reportWs = ThisWorkbook.Sheets("Machine Times Report")
    GoTo SheetFound

CreateReportSheet:
    ' If we couldn't find it, create it
    Err.Clear
    Set reportWs = ThisWorkbook.Sheets.Add
    reportWs.Name = "Machine Times Report"

SheetFound:
    ' Clear old data
    reportWs.Cells.Clear
    
    ' Add headers
    reportWs.Cells(1, 1).Value = "Machine Category"
    reportWs.Cells(1, 2).Value = "Min Time (hh:mm:ss)"
    reportWs.Cells(1, 3).Value = "Max Time (hh:mm:ss)"
    reportWs.Cells(1, 4).Value = "Sheet Name"
    reportRow = 2
    
    '---------------------------------------------
    ' 3) Loop each Worksheet
    '---------------------------------------------
    Dim anySheetsFound As Boolean
    anySheetsFound = False

    For Each ws In ThisWorkbook.Sheets
        ' Make sure ws is a real Worksheet (not a Chart or dialog sheet)
        If ws.Type = xlWorksheet Then
            
            Debug.Print "Checking sheet: " & ws.Name
            If ws.Name Like "*_BBM_Export_Timings" Then
                anySheetsFound = True
                Debug.Print " -> Matched pattern: " & ws.Name
                
                '---------------------------------------------
                ' 4) Define columns to scan
                '---------------------------------------------
                colStart = ws.Columns("H").Column  ' Start at column H
                colEnd = ws.Columns("BQ").Column   ' End at column BQ
                
                ' Create a new dictionary for times
                Set machineTimes = CreateObject("Scripting.Dictionary")
                For Each category In machineCategories.Keys
                    machineTimes.Add category, New Collection
                Next category
                
                '---------------------------------------------
                ' 5) Loop columns H through BQ
                '---------------------------------------------
                For i = colStart To colEnd
                    
                    machineName = ws.Cells(1, i).Value
                    timeValue = ws.Cells(14, i).Value
                    
                    ' Debugging info:
                    Debug.Print "   Column: " & i & _
                                " | Machine Name: " & machineName & _
                                " (" & TypeName(machineName) & ")" & _
                                " | Time Value: " & timeValue & _
                                " (" & TypeName(timeValue) & ")"
                    
                    ' a) Skip if machine name is an error or empty
                    If IsError(machineName) Or IsEmpty(machineName) Then
                        Debug.Print "   -> Skipped (empty/error machine name)"
                        GoTo SkipColumn
                    End If
                    
                    ' b) Extract base machine name
                    cleanMachineName = ExtractMachineName(CStr(machineName))
                    If cleanMachineName = "" Then
                        Debug.Print "   -> Skipped (invalid machine name prefix)"
                        GoTo SkipColumn
                    End If
                    
                    ' c) Skip if timeValue is error or empty
                    If IsError(timeValue) Or IsEmpty(timeValue) Then
                        Debug.Print "   -> Skipped (time is #Error or empty)"
                        GoTo SkipColumn
                    End If
                    
                    ' d) Confirm that timeValue is recognized as a date/time
                    If Not IsDate(timeValue) Then
                        Debug.Print "   -> Skipped (Not a valid date/time)"
                        GoTo SkipColumn
                    End If
                    
                    ' e) Convert timeValue to a Double
                    Debug.Print "   -> Attempting CDbl conversion..."
                    numericTime = CDbl(CDate(timeValue))
                    
                    ' Optional sanity check: if numericTime = 0 but wasn't "0"
                    ' (this might happen if the cell is text or date 0 = 12/31/1899)
                    ' We'll just let 0 pass if it truly is 12/31/1899 00:00:00
                    Debug.Print "   -> numericTime = "; numericTime
                    
                    ' f) Match category & store the numeric time
                    For Each category In machineCategories.Keys
                        If cleanMachineName = machineCategories(category) Then
                            machineTimes(category).Add numericTime
                            Exit For
                        End If
                    Next category
                    
SkipColumn:
                Next i  ' Next column
                
                '---------------------------------------------
                ' 6) For each category, compute Min/Max
                '---------------------------------------------
                For Each category In machineCategories.Keys
                    If machineTimes(category).Count > 0 Then
                        ' Convert the collection to a Double array
                        ReDim timeArray(1 To machineTimes(category).Count)
                        For timeIndex = 1 To machineTimes(category).Count
                            timeArray(timeIndex) = machineTimes(category).Item(timeIndex)
                        Next timeIndex
                        
                        ' Get min & max
                        minTime = Application.Min(timeArray)
                        maxTime = Application.Max(timeArray)
                        
                        ' Write to report
                        reportWs.Cells(reportRow, 1).Value = category
                        reportWs.Cells(reportRow, 2).Value = Format(minTime, "hh:mm:ss")
                        reportWs.Cells(reportRow, 3).Value = Format(maxTime, "hh:mm:ss")
                        reportWs.Cells(reportRow, 4).Value = ws.Name
                        reportRow = reportRow + 1
                    End If
                Next category
                
            End If ' If name Like ...
        End If ' If ws.Type = xlWorksheet
    Next ws
    
    '---------------------------------------------
    ' 7) Format the report
    '---------------------------------------------
    If anySheetsFound Then
        reportWs.Columns.AutoFit
        MsgBox "Processing complete. Check 'Machine Times Report'.", vbInformation
    Else
        MsgBox "No sheets matched the '*_BBM_Export_Timings' pattern.", vbInformation
    End If
    
    Exit Sub
    
End Sub

'================================================================
' Extracts the base machine name (e.g., "VIRTIXEN" from "VIRTIXEN630S")
'================================================================
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
