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
    reportWs.Range("A1").Value = "Machine Category"
    reportWs.Range("B1").Value = "Min Time (hh:mm:ss)"
    reportWs.Range("C1").Value = "Max Time (hh:mm:ss)"
    reportWs.Range("D1").Value = "Sheet Name"
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
            ' 4) Columns H to BQ → col 8 to col 69
            '    Row 1 → machine name
            '    Row 14 → time value
            '---------------------------------------------
            For i = 8 To 69  ' 8=H, 69=BQ
                machineName = ws.Cells(1, i).Value
                timeValue = ws.Cells(14, i).Value
                
                 Debug.Print "Column " & i & _
                " | MachineName='" & thisName & "'" & _
                " => Extracted='" & ExtractMachineName(thisName) & "'" & _
                " | Row14='" & thisTime & "'"
                
                ' First, skip if machineName is empty/error
                If Not IsError(machineName) And Not IsEmpty(machineName) Then
                    ' Try to extract the base name (e.g. "VIRTPPC" from "VIRTPPC675S")
                    cleanMachineName = ExtractMachineName(CStr(machineName))
                    
                    If Len(cleanMachineName) > 0 Then
                        ' Next, skip if timeValue is empty/error
                        If Not IsError(timeValue) And Not IsEmpty(timeValue) Then
                            
                            ' Attempt to interpret it as a Date/Time
                            If IsDate(timeValue) Then
                                ' If Excel already sees it as a date/time
                                numericTime = CDbl(CDate(timeValue))
                            Else
                                ' If "IsDate" fails, try TimeValue(...) in case it's text like "00:04:19"
                                On Error Resume Next
                                numericTime = CDbl(TimeValue(CStr(timeValue)))
                                On Error GoTo 0
                            End If
                            
                            ' If numericTime > 0, we assume we parsed a valid time
                            If numericTime > 0 Then
                                ' Add it to the proper category
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
            ' 5) For each category, compute Min/Max
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
            
        End If
    Next ws
    
    '---------------------------------------------
    ' 6) Format the report
    '---------------------------------------------
    reportWs.Columns.AutoFit
    MsgBox "Processing complete. Check the 'Machine Times Report' sheet."
    
End Sub


'================================================
' ExtractMachineName: e.g. "VIRTPPC" from "VIRTPPC675S"
'================================================
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
