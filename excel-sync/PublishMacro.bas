Attribute VB_Name = "PublishMacro"
Option Explicit

' Core row-publishing logic for the Holly Hill Surplus shop sync, plus a
' manual "force resync this row" entry point.
'
' PublishRow is the shared engine: it builds pending_lot.json for one row,
' runs publish_lot.py (waiting for it to finish), reads result.json, and
' writes the outcome back into columns H (Last Synced) and I (Sync Status).
' It is silent (no MsgBox) and returns True/False so it can be called from
' both PublishSelectedRow (this module) and Workbook_BeforeSave (in the
' ThisWorkbook module) without popping up a dialog per row during auto-sync.
'
' Expects this workbook to live in the same excel-sync\ folder as
' publish_lot.py, config.ini, and requirements.txt.

Sub PublishSelectedRow()
    ' Manual "force resync this row" button. Re-sends every photo path
    ' currently in column G (not just new ones) and never clears images --
    ' it's a deliberate full resync, distinct from the auto-sync-on-save
    ' incremental behavior in Workbook_BeforeSave.
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim r As Long
    r = ActiveCell.Row

    If r <= 1 Then
        MsgBox "Please click into a data row (not the header row) before running Publish.", _
            vbExclamation, "Publish Lot"
        Exit Sub
    End If

    Dim lotID As String
    lotID = Trim(ws.Cells(r, 1).Value)

    ' Friendly pre-checks with specific messages -- PublishRow itself fails
    ' silently (by design, so auto-sync doesn't pop up a dialog per row).
    If lotID = "" Then
        MsgBox "Lot # / SKU is required (column A).", vbExclamation, "Publish Lot"
        Exit Sub
    End If
    If Trim(ws.Cells(r, 3).Value) = "" Then
        MsgBox "Item Name is required (column C).", vbExclamation, "Publish Lot"
        Exit Sub
    End If
    If Not IsNumeric(ws.Cells(r, 5).Value) Then
        MsgBox "Price ($) must be a number (column E).", vbExclamation, "Publish Lot"
        Exit Sub
    End If

    Dim statusVal As String
    statusVal = Trim(ws.Cells(r, 6).Value)
    Dim validStatuses As Variant
    validStatuses = Array("available", "reserved", "sold", "shipped")
    Dim isValidStatus As Boolean
    isValidStatus = False
    Dim i As Long
    For i = LBound(validStatuses) To UBound(validStatuses)
        If LCase(statusVal) = validStatuses(i) Then isValidStatus = True
    Next i
    If Not isValidStatus Then
        MsgBox "Status (column F) must be one of: available, reserved, sold, shipped.", _
            vbExclamation, "Publish Lot"
        Exit Sub
    End If

    Dim photoPathsCSV As String
    photoPathsCSV = Trim(ws.Cells(r, 7).Value)

    Application.Cursor = xlWait
    Dim success As Boolean
    success = PublishRow(ws, r, photoPathsCSV, False)
    Application.Cursor = xlDefault

    If success Then
        ' Keep the fingerprint / last-synced-photos columns current too, so
        ' a routine save right after this doesn't think there's still an
        ' unsynced change pending for this row.
        ws.Cells(r, 10).Value = BuildFingerprint(ws, r)
        ws.Cells(r, 11).Value = photoPathsCSV
        MsgBox "Lot '" & lotID & "' published successfully.", vbInformation, "Publish Lot"
    Else
        MsgBox "Publish failed for lot '" & lotID & "'. See column I (Sync Status) for details.", _
            vbCritical, "Publish Lot"
    End If
End Sub

' Publishes one row. photoPathsCSV is the semicolon-separated list of photo
' paths to send as this call's photo_paths (the caller decides whether
' that's "all current paths" -- PublishSelectedRow -- or just the new ones
' -- Workbook_BeforeSave's incremental sync). If clearImages is True, the
' sync-lot endpoint is told to replace image_urls with just what's sent
' here (photoPathsCSV should be "" in that case) instead of appending.
'
' Silent by design -- no MsgBox -- so it can run unattended from
' Workbook_BeforeSave across many rows without interrupting the save.
Function PublishRow(ws As Worksheet, rowNum As Long, photoPathsCSV As String, clearImages As Boolean) As Boolean
    Dim lotID As String, category As String, itemName As String
    Dim description As String, statusVal As String
    Dim priceCellValue As Variant

    lotID = Trim(ws.Cells(rowNum, 1).Value)
    category = Trim(ws.Cells(rowNum, 2).Value)
    itemName = Trim(ws.Cells(rowNum, 3).Value)
    description = Trim(ws.Cells(rowNum, 4).Value)
    priceCellValue = ws.Cells(rowNum, 5).Value
    statusVal = Trim(ws.Cells(rowNum, 6).Value)

    ' --- Validation (silent -- just fail) --------------------------------

    If lotID = "" Or itemName = "" Then
        PublishRow = False
        Exit Function
    End If

    If Not IsNumeric(priceCellValue) Then
        PublishRow = False
        Exit Function
    End If
    Dim priceVal As Double
    priceVal = CDbl(priceCellValue)

    Dim validStatuses As Variant
    validStatuses = Array("available", "reserved", "sold", "shipped")
    Dim isValidStatus As Boolean
    isValidStatus = False
    Dim i As Long
    For i = LBound(validStatuses) To UBound(validStatuses)
        If LCase(statusVal) = validStatuses(i) Then
            statusVal = validStatuses(i) ' normalize case
            isValidStatus = True
            Exit For
        End If
    Next i
    If Not isValidStatus Then
        PublishRow = False
        Exit Function
    End If

    ' --- Build the JSON payload -------------------------------------------

    Dim priceStr As String
    priceStr = Replace(CStr(priceVal), ",", ".")

    Dim jsonPayload As String
    jsonPayload = "{" & vbCrLf
    jsonPayload = jsonPayload & "  ""id"": """ & JSONEscape(lotID) & """," & vbCrLf
    jsonPayload = jsonPayload & "  ""category"": """ & JSONEscape(category) & """," & vbCrLf
    jsonPayload = jsonPayload & "  ""name"": """ & JSONEscape(itemName) & """," & vbCrLf
    jsonPayload = jsonPayload & "  ""description"": """ & JSONEscape(description) & """," & vbCrLf
    jsonPayload = jsonPayload & "  ""price_dollars"": " & priceStr & "," & vbCrLf
    jsonPayload = jsonPayload & "  ""status"": """ & statusVal & """," & vbCrLf
    If clearImages Then
        jsonPayload = jsonPayload & "  ""clear_images"": true," & vbCrLf
    End If
    jsonPayload = jsonPayload & "  ""photo_paths"": " & BuildPhotoPathsJSON(photoPathsCSV) & vbCrLf
    jsonPayload = jsonPayload & "}"

    ' --- Write pending_lot.json (UTF-8, no BOM) ---------------------------

    Dim pendingPath As String
    pendingPath = ThisWorkbook.Path & "\pending_lot.json"
    WriteUTF8File pendingPath, jsonPayload

    ' --- Run the sync script, waiting for it to finish ---------------------

    Dim shellObj As Object
    Set shellObj = CreateObject("WScript.Shell")

    Dim cmd As String
    cmd = "cmd.exe /c cd /d " & Chr(34) & ThisWorkbook.Path & Chr(34) & _
        " && python publish_lot.py"

    Dim exitCode As Long
    exitCode = shellObj.Run(cmd, 0, True) ' 0 = hidden window, True = wait

    ' --- Read result.json and update H / I for this row --------------------

    Dim resultPath As String
    resultPath = ThisWorkbook.Path & "\result.json"

    ws.Cells(rowNum, 8).Value = Now ' Last Synced

    If Dir(resultPath) = "" Then
        ws.Cells(rowNum, 9).Value = "ERROR: result.json was not created -- check that Python " & _
            "is installed and on PATH"
        PublishRow = False
        Exit Function
    End If

    Dim resultText As String
    resultText = ReadTextFile(resultPath)

    Dim isSuccess As Boolean
    isSuccess = (InStr(1, resultText, Chr(34) & "success" & Chr(34) & ": true") > 0)

    If isSuccess Then
        ws.Cells(rowNum, 9).Value = "OK"
        PublishRow = True
    Else
        Dim errMsg As String
        errMsg = ExtractJSONStringValue(resultText, "error")
        If errMsg = "" Then errMsg = "Unknown error -- check result.json manually."
        ws.Cells(rowNum, 9).Value = "ERROR: " & errMsg
        PublishRow = False
    End If
End Function

' A-G concatenated with "|" separators -- cheap change-detection fingerprint
' stored in column J after a successful sync and compared against on the
' next save.
Function BuildFingerprint(ws As Worksheet, r As Long) As String
    BuildFingerprint = CStr(ws.Cells(r, 1).Value) & "|" & _
        CStr(ws.Cells(r, 2).Value) & "|" & _
        CStr(ws.Cells(r, 3).Value) & "|" & _
        CStr(ws.Cells(r, 4).Value) & "|" & _
        CStr(ws.Cells(r, 5).Value) & "|" & _
        CStr(ws.Cells(r, 6).Value) & "|" & _
        CStr(ws.Cells(r, 7).Value)
End Function

' Returns the semicolon-separated paths present in currentCSV but not in
' previousCSV (case-insensitive, since Windows file paths are), preserving
' currentCSV's order. Used by Workbook_BeforeSave to only send genuinely
' new photos rather than re-uploading ones already synced.
Function GetNewPaths(ByVal currentCSV As String, ByVal previousCSV As String) As String
    currentCSV = Trim(currentCSV)
    If currentCSV = "" Then
        GetNewPaths = ""
        Exit Function
    End If

    Dim currentParts() As String
    Dim previousParts() As String
    currentParts = Split(currentCSV, ";")
    previousParts = Split(previousCSV, ";")

    Dim result As String
    Dim firstItem As Boolean
    firstItem = True

    Dim i As Long, j As Long
    Dim onePath As String
    Dim alreadySynced As Boolean

    For i = LBound(currentParts) To UBound(currentParts)
        onePath = Trim(currentParts(i))
        If onePath <> "" Then
            alreadySynced = False
            For j = LBound(previousParts) To UBound(previousParts)
                If StrComp(Trim(previousParts(j)), onePath, vbTextCompare) = 0 Then
                    alreadySynced = True
                    Exit For
                End If
            Next j
            If Not alreadySynced Then
                If Not firstItem Then result = result & ";"
                result = result & onePath
                firstItem = False
            End If
        End If
    Next i

    GetNewPaths = result
End Function

' --- Helpers ----------------------------------------------------------------

Private Function BuildPhotoPathsJSON(ByVal photoPathsCSV As String) As String
    Dim result As String
    result = "["
    photoPathsCSV = Trim(photoPathsCSV)
    If photoPathsCSV <> "" Then
        Dim parts() As String
        parts = Split(photoPathsCSV, ";")
        Dim p As Long
        Dim firstItem As Boolean
        firstItem = True
        For p = LBound(parts) To UBound(parts)
            Dim onePath As String
            onePath = Trim(parts(p))
            If onePath <> "" Then
                If Not firstItem Then result = result & ", "
                result = result & Chr(34) & JSONEscape(onePath) & Chr(34)
                firstItem = False
            End If
        Next p
    End If
    result = result & "]"
    BuildPhotoPathsJSON = result
End Function

Private Function JSONEscape(ByVal text As String) As String
    Dim result As String
    result = text
    result = Replace(result, "\", "\\")
    result = Replace(result, Chr(34), "\" & Chr(34))
    result = Replace(result, vbCrLf, "\n")
    result = Replace(result, vbCr, "\n")
    result = Replace(result, vbLf, "\n")
    result = Replace(result, vbTab, "\t")
    JSONEscape = result
End Function

Private Sub WriteUTF8File(ByVal filePath As String, ByVal content As String)
    ' ADODB.Stream writes a UTF-8 BOM by default -- strip it so the file
    ' is plain UTF-8, matching what Python's open(..., encoding="utf-8")
    ' expects.
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' adTypeText
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText content

    stream.Position = 0
    stream.Type = 1 ' adTypeBinary
    stream.Position = 3 ' skip the 3-byte BOM

    Dim noBOM As Object
    Set noBOM = CreateObject("ADODB.Stream")
    noBOM.Type = 1 ' adTypeBinary
    noBOM.Open
    stream.CopyTo noBOM
    noBOM.SaveToFile filePath, 2 ' adSaveCreateOverWrite
    noBOM.Close
    stream.Close
End Sub

Private Function ReadTextFile(ByVal filePath As String) As String
    Dim fileNum As Integer
    fileNum = FreeFile
    Dim content As String
    Dim oneLine As String
    Open filePath For Input As #fileNum
    Do While Not EOF(fileNum)
        Line Input #fileNum, oneLine
        content = content & oneLine & vbCrLf
    Loop
    Close #fileNum
    ReadTextFile = content
End Function

' Extracts the string value for "key": "..." out of a JSON text blob,
' respecting \" and \\ escapes. Good enough for reading back our own
' result.json without needing a full JSON parser library.
Private Function ExtractJSONStringValue(ByVal jsonText As String, ByVal key As String) As String
    Dim searchKey As String
    searchKey = Chr(34) & key & Chr(34) & ":"

    Dim startPos As Long
    startPos = InStr(1, jsonText, searchKey)
    If startPos = 0 Then
        ExtractJSONStringValue = ""
        Exit Function
    End If

    Dim pos As Long
    pos = startPos + Len(searchKey)

    Do While pos <= Len(jsonText) And (Mid(jsonText, pos, 1) = " " Or Mid(jsonText, pos, 1) = vbTab)
        pos = pos + 1
    Loop

    If Mid(jsonText, pos, 1) <> Chr(34) Then
        ExtractJSONStringValue = ""
        Exit Function
    End If
    pos = pos + 1 ' move past opening quote

    Dim result As String
    Dim ch As String
    Do While pos <= Len(jsonText)
        ch = Mid(jsonText, pos, 1)
        If ch = "\" Then
            Dim nextCh As String
            nextCh = Mid(jsonText, pos + 1, 1)
            Select Case nextCh
                Case Chr(34): result = result & Chr(34)
                Case "\": result = result & "\"
                Case "n": result = result & vbLf
                Case "r": result = result & vbCr
                Case "t": result = result & vbTab
                Case "/": result = result & "/"
                Case Else: result = result & nextCh
            End Select
            pos = pos + 2
        ElseIf ch = Chr(34) Then
            Exit Do
        Else
            result = result & ch
            pos = pos + 1
        End If
    Loop

    ExtractJSONStringValue = result
End Function
