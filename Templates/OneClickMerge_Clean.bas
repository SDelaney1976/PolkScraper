
Attribute VB_Name = "OneClickMerge_Clean"
Option Explicit

' ============================
' ONE-CLICK MERGE (CLEAN SLATE)
' ============================
' You get two safe entry points:
'   1) Run_Merge_From_CSV      -> Recommended on Mac. No Excel automation, never freezes.
'   2) Run_Merge_From_Excel    -> Tries Excel automation; if it fails, use CSV method.
'
' Your Word template should contain merge fields:
'   «Name», «Address_1», «City», «State», «Zip», «Case_Number»
'
' Your data file (CSV or Excel) must include columns (any of the accepted aliases):
'   Name: Name, RecipientFullName, Defendant Name, Defendant_Name
'   Address_1: Address_1, Address 1, Address1, Street, Street Address
'   City: City, City/Town
'   State: State, ST, State Abbrev
'   Zip: Zip, ZIP, Zip Code, ZipCode, Postal, Postal Code
'   Case_Number: Case_Number, Case Number, CaseNumber, Case No, CaseNo
'   Capture Date filter: Capture Date, Capture_Date, CaptureDate, Captured Date, Capture Dt
'
' Both flows prompt:
'   - Pick your data file
'   - Enter a date (defaults to today)
'   - Choose: Merge to NEW DOCUMENT (review) or PRINT directly

Public Sub Run_Merge_From_CSV()
    On Error GoTo ErrH
    Dim csvPath As String
    csvPath = PickFile(False) ' False => CSV
    If Len(csvPath) = 0 Then Exit Sub

    Dim targetDate As Date
    If Not PromptForDate(targetDate) Then Exit Sub

    ' Read CSV, find header indices, write filtered temp CSV with canonical headers
    Dim tmpCsv As String
    tmpCsv = FilterCsvByDate(csvPath, targetDate)
    If Len(tmpCsv) = 0 Then
        MsgBox "No rows matched the selected date.", vbInformation, "Mail Merge"
        Exit Sub
    End If

    Dim sendToPrinter As VbMsgBoxResult
    sendToPrinter = MsgBox("Matched rows for " & Format$(targetDate, "m/d/yyyy") & "." & vbCrLf & _
                           "Click YES to PRINT, NO to merge to a NEW DOCUMENT for review.", _
                           vbYesNoCancel + vbQuestion, "Mail Merge")
    If sendToPrinter = vbCancel Then GoTo Cleanup

    ExecuteMerge tmpCsv, (sendToPrinter = vbYes)

Cleanup:
    On Error Resume Next
    Kill tmpCsv
    Exit Sub
ErrH:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Run_Merge_From_CSV"
End Sub

Public Sub Run_Merge_From_Excel()
    On Error GoTo ErrH
    Dim xlsPath As String
    xlsPath = PickFile(True) ' True => Excel
    If Len(xlsPath) = 0 Then Exit Sub

    Dim targetDate As Date
    If Not PromptForDate(targetDate) Then Exit Sub

    Dim tmpCsv As String
    tmpCsv = BuildCsvFilteredByDate_Excel(xlsPath, targetDate)
    If Len(tmpCsv) = 0 Then
        MsgBox "No rows matched the selected date, or Excel couldn't be automated." & vbCrLf & _
               "Tip: Save your workbook as CSV and run 'Run_Merge_From_CSV'.", vbInformation, "Mail Merge"
        Exit Sub
    End If

    Dim sendToPrinter As VbMsgBoxResult
    sendToPrinter = MsgBox("Matched rows for " & Format$(targetDate, "m/d/yyyy") & "." & vbCrLf & _
                           "Click YES to PRINT, NO to merge to a NEW DOCUMENT for review.", _
                           vbYesNoCancel + vbQuestion, "Mail Merge")
    If sendToPrinter = vbCancel Then GoTo Cleanup

    ExecuteMerge tmpCsv, (sendToPrinter = vbYes)

Cleanup:
    On Error Resume Next
    Kill tmpCsv
    Exit Sub
ErrH:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Run_Merge_From_Excel"
End Sub

' ---------------- Core merge ----------------
Private Sub ExecuteMerge(ByVal dataPath As String, ByVal toPrinter As Boolean)
    With ActiveDocument.MailMerge
        .MainDocumentType = wdFormLetters
        .OpenDataSource Name:=dataPath, ConfirmConversions:=False, ReadOnly:=True, _
                        AddToRecentFiles:=False, Revert:=False, Format:=wdOpenFormatAuto
        If toPrinter Then
            .Destination = wdSendToPrinter
        Else
            .Destination = wdSendToNewDocument
        End If
        .SuppressBlankLines = True
        .Execute Pause:=False
    End With
End Sub

' --------------- File pickers ---------------
' isExcel = True  -> Excel files only
' isExcel = False -> CSV files only
Private Function PickFile(ByVal isExcel As Boolean) As String
    Dim path As String
    path = ""

    ' Try FileDialog first (works on Windows; may fail on Mac)
    On Error Resume Next
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    If Err.Number = 0 Then
        With fd
            .Title = IIf(isExcel, "Select your Excel source file", "Select your CSV source file")
            .Filters.Clear
            If isExcel Then
                .Filters.Add "Excel Files", "*.xlsx;*.xlsm;*.xls"
            Else
                .Filters.Add "CSV Files", "*.csv"
            End If
            .AllowMultiSelect = False
            If .Show = -1 Then path = .SelectedItems(1)
        End With
    End If
    On Error GoTo 0

    If Len(path) > 0 Then
        PickFile = path
        Exit Function
    End If

    ' Mac-safe fallback via AppleScript (Word Mac supports MacScript in VBA)
    #If Mac Then
        On Error Resume Next
        If isExcel Then
            path = MacScript("choose file with prompt ""Select your Excel file"" of type {""XLSX"",""XLSM"",""XLS""} as alias" & vbCr & "POSIX path of result")
        Else
            path = MacScript("choose file with prompt ""Select your CSV file"" of type {""CSV""} as alias" & vbCr & "POSIX path of result")
        End If
        On Error GoTo 0
        If Len(path) > 0 Then
            PickFile = path
            Exit Function
        End If
    #End If

    ' Final fallback (manual): use Word's FileOpen dialog (may open doc; cancel after noting path)
    ' Recommend CSV flow if all else fails.
    PickFile = ""
End Function

' --------------- Date prompt ---------------
Private Function PromptForDate(ByRef outDate As Date) As Boolean
    Dim s As String
    s = InputBox("Enter the Capture Date to print (e.g., 8/14/2025). Leave blank to cancel.", _
                 "Capture Date", Format$(Date, "m/d/yyyy"))
    If Len(Trim$(s)) = 0 Then
        PromptForDate = False
        Exit Function
    End If
    If IsDate(s) Then
        outDate = DateValue(CDate(s))
        PromptForDate = True
    Else
        MsgBox "That doesn't look like a valid date. Try again (e.g., 8/14/2025).", vbExclamation, "Invalid Date"
        PromptForDate = PromptForDate(outDate)
    End If
End Function

' --------------- CSV (no Excel) path ---------------
' Reads a user-supplied CSV, finds accepted headers, filters rows by date,
' writes canonical header CSV for merge.
Private Function FilterCsvByDate(ByVal csvPath As String, ByVal targetDate As Date) As String
    Dim f As Integer, line As String, headers() As String
    Dim idxName As Long, idxAddr1 As Long, idxCity As Long, idxState As Long, idxZip As Long, idxCase As Long, idxCapture As Long
    Dim tmpPath As String, outF As Integer
    Dim wroteAny As Boolean

    ' Open and read header
    f = FreeFile(0)
    Open csvPath For Input As #f
    If EOF(f) Then Close #f: Exit Function
    Line Input #f, line
    headers = SplitCsv(line)

    idxName = FindHeader(headers, Array("Name", "RecipientFullName", "Defendant Name", "Defendant_Name"))
    idxAddr1 = FindHeader(headers, Array("Address_1", "Address 1", "Address1", "Street", "Street Address"))
    idxCity = FindHeader(headers, Array("City", "City/Town"))
    idxState = FindHeader(headers, Array("State", "ST", "State Abbrev"))
    idxZip = FindHeader(headers, Array("Zip", "ZIP", "Zip Code", "ZipCode", "Postal", "Postal Code"))
    idxCase = FindHeader(headers, Array("Case_Number", "Case Number", "CaseNumber", "Case No", "CaseNo"))
    idxCapture = FindHeader(headers, Array("Capture Date", "Capture_Date", "CaptureDate", "Captured Date", "Capture Dt"))

    If idxCapture = -1 Then
        MsgBox "Couldn't find a 'Capture Date' column in your CSV.", vbCritical, "Mail Merge"
        Close #f
        Exit Function
    End If

    tmpPath = Environ$("TEMP") & "\merge_bydate_" & Format$(Now, "yyyymmdd_hhnnss") & ".csv"
    outF = FreeFile(0)
    Open tmpPath For Output As #outF
    Print #outF, "Name,Address_1,City,State,Zip,Case_Number"

    Do While Not EOF(f)
        Line Input #f, line
        Dim cells() As String
        cells = SplitCsv(line)

        If UBound(cells) >= idxCapture Then
            Dim v As String: v = cells(idxCapture)
            Dim d As Date
            If TryParseDate(v, d) Then
                If DateValue(d) = targetDate Then
                    Print #outF, CsvEscape(GetCell(cells, idxName)) & "," & _
                                 CsvEscape(GetCell(cells, idxAddr1)) & "," & _
                                 CsvEscape(GetCell(cells, idxCity)) & "," & _
                                 CsvEscape(GetCell(cells, idxState)) & "," & _
                                 CsvEscape(GetCell(cells, idxZip)) & "," & _
                                 CsvEscape(GetCell(cells, idxCase))
                    wroteAny = True
                End If
            End If
        End If
    Loop
    Close #f
    Close #outF

    If wroteAny Then
        FilterCsvByDate = tmpPath
    Else
        Kill tmpPath
        FilterCsvByDate = ""
    End If
End Function

Private Function GetCell(ByRef arr() As String, ByVal idx As Long) As String
    If idx >= 0 And idx <= UBound(arr) Then
        GetCell = arr(idx)
    Else
        GetCell = ""
    End If
End Function

Private Function TryParseDate(ByVal s As String, ByRef outDate As Date) As Boolean
    On Error GoTo Bad
    If IsDate(s) Then
        outDate = DateValue(CDate(s))
        TryParseDate = True
        Exit Function
    End If
Bad:
    TryParseDate = False
End Function

' ---------------- Excel (automation) path ----------------
Private Function BuildCsvFilteredByDate_Excel(ByVal xlsPath As String, ByVal targetDate As Date) As String
    On Error GoTo ExcelErr
    Dim xlApp As Object, wb As Object, ws As Object
    Dim usedRows As Long, usedCols As Long, r As Long, c As Long
    Dim headers() As String
    Dim idxName As Long, idxAddr1 As Long, idxCity As Long, idxState As Long, idxZip As Long, idxCase As Long, idxCapture As Long
    Dim tmpPath As String, f As Integer, line As String
    Dim wroteAny As Boolean

    Dim aliasName As Variant, aliasAddr1 As Variant, aliasCity As Variant, aliasState As Variant, aliasZip As Variant, aliasCase As Variant, aliasCapDate As Variant
    aliasName = Array("Name", "RecipientFullName", "Defendant Name", "Defendant_Name")
    aliasAddr1 = Array("Address_1", "Address 1", "Address1", "Street", "Street Address")
    aliasCity = Array("City", "City/Town")
    aliasState = Array("State", "ST", "State Abbrev")
    aliasZip = Array("Zip", "ZIP", "Zip Code", "ZipCode", "Postal", "Postal Code")
    aliasCase = Array("Case_Number", "Case Number", "CaseNumber", "Case No", "CaseNo")
    aliasCapDate = Array("Capture Date", "Capture_Date", "CaptureDate", "Captured Date", "Capture Dt")

    Set xlApp = CreateObject("Excel.Application")
    xlApp.ScreenUpdating = False
    xlApp.DisplayAlerts = False

    Set wb = xlApp.Workbooks.Open(Filename:=xlsPath, ReadOnly:=True)
    If wb.Worksheets.Count = 1 Then
        Set ws = wb.Worksheets(1)
    Else
        Set ws = wb.ActiveSheet
    End If

    usedRows = ws.UsedRange.Rows.Count
    usedCols = ws.UsedRange.Columns.Count
    If usedRows < 2 Then GoTo NoRows

    ReDim headers(1 To usedCols)
    For c = 1 To usedCols
        headers(c) = Trim$(CStr(ws.Cells(1, c).Value))
    Next c

    idxName = FindHeader(headers, aliasName)
    idxAddr1 = FindHeader(headers, aliasAddr1)
    idxCity = FindHeader(headers, aliasCity)
    idxState = FindHeader(headers, aliasState)
    idxZip = FindHeader(headers, aliasZip)
    idxCase = FindHeader(headers, aliasCase)
    idxCapture = FindHeader(headers, aliasCapDate)

    If idxCapture = -1 Then GoTo Cleanup

    tmpPath = Environ$("TEMP") & "\merge_bydate_" & Format$(Now, "yyyymmdd_hhnnss") & ".csv"
    f = FreeFile(0)
    Open tmpPath For Output As #f
    Print #f, "Name,Address_1,City,State,Zip,Case_Number"

    Dim v As Variant, maybe As Date
    For r = 2 To usedRows
        v = ws.Cells(r, idxCapture).Value
        If IsDate(v) Then
            If DateValue(v) = targetDate Then
                Print #f, CsvEscape(ws.Cells(r, idxName).Text) & "," & _
                          CsvEscape(ws.Cells(r, idxAddr1).Text) & "," & _
                          CsvEscape(ws.Cells(r, idxCity).Text) & "," & _
                          CsvEscape(ws.Cells(r, idxState).Text) & "," & _
                          CsvEscape(ws.Cells(r, idxZip).Text) & "," & _
                          CsvEscape(ws.Cells(r, idxCase).Text)
                wroteAny = True
            End If
        ElseIf Len(Trim$(CStr(v))) > 0 Then
            On Error Resume Next
            maybe = CDate(CStr(v))
            If Err.Number = 0 Then
                If DateValue(maybe) = targetDate Then
                    Print #f, CsvEscape(ws.Cells(r, idxName).Text) & "," & _
                              CsvEscape(ws.Cells(r, idxAddr1).Text) & "," & _
                              CsvEscape(ws.Cells(r, idxCity).Text) & "," & _
                              CsvEscape(ws.Cells(r, idxState).Text) & "," & _
                              CsvEscape(ws.Cells(r, idxZip).Text) & "," & _
                              CsvEscape(ws.Cells(r, idxCase).Text)
                    wroteAny = True
                End If
            End If
            Err.Clear
            On Error GoTo 0
        End If
    Next r
    Close #f

    If wroteAny Then
        BuildCsvFilteredByDate_Excel = tmpPath
    Else
        Kill tmpPath
        BuildCsvFilteredByDate_Excel = ""
    End If

Cleanup:
    On Error Resume Next
    wb.Close SaveChanges:=False
    xlApp.Quit
    Set ws = Nothing: Set wb = Nothing: Set xlApp = Nothing
    Exit Function
NoRows:
    GoTo Cleanup
ExcelErr:
    BuildCsvFilteredByDate_Excel = ""
    Resume Cleanup
End Function

' ---------------- Helpers ----------------
' Find header by list of aliases (case-insensitive). Returns 0-based index for CSV split;
' returns 1-based index for Excel headers array. To normalize, we return 0-based for CSV
' and 1-based for Excel, but we handle that in callers.
Private Function FindHeader(ByVal headers As Variant, ByVal aliases As Variant) As Long
    Dim i As Long, j As Long
    For j = LBound(aliases) To UBound(aliases)
        For i = LBound(headers) To UBound(headers)
            If StrComp(CStr(headers(i)), CStr(aliases(j)), vbTextCompare) = 0 Then
                FindHeader = i
                Exit Function
            End If
        Next i
    Next j
    FindHeader = IIf(IsArray(headers), -1, 0)
End Function

' Split a CSV line into array (handles simple quoted commas)
Private Function SplitCsv(ByVal s As String) As String()
    Dim result() As String
    Dim i As Long, inQuotes As Boolean, ch As String, cur As String
    ReDim result(0 To 0)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch = """" Then
            inQuotes = Not inQuotes
            cur = cur & ch
        ElseIf ch = "," And Not inQuotes Then
            result(UBound(result)) = cur
            ReDim Preserve result(0 To UBound(result) + 1)
            cur = ""
        Else
            cur = cur & ch
        End If
    Next i
    result(UBound(result)) = cur
    SplitCsv = result
End Function

Private Function CsvEscape(ByVal s As String) As String
    Dim t As String
    t = Replace(s, """", """""")
    CsvEscape = """" & t & """"
End Function
