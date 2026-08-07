Attribute VB_Name = "SoldLotsMacro"
Option Explicit

' Read-only "Sold Lots" tracker for Holly Hill Surplus. Pulls order + lot
' data from the website via pull_sold_orders.py and keeps this sheet
' current. This workbook NEVER writes anything back to the website -- it
' is a one-way, pull-only mirror, entirely separate from lots.xlsm.
'
' Columns (Sold Lots sheet):
'   A  Order ID       -- server-sourced, used as the match key on refresh
'   B  Date Sold
'   C  Lot #/SKU
'   D  Item Name
'   E  Category
'   F  Sale Price
'   G  Buyer Email
'   H  Order Status
'   I  Fulfilled      -- LOCAL ONLY, manual Y/N, never touched by refresh
'   J  Notes          -- LOCAL ONLY, freeform, never touched by refresh
'
' RefreshSoldLots is the shared engine: for each order the server returns,
' it either appends a new row (A:H filled, I:J left blank) or -- if that
' Order ID is already present somewhere on the sheet -- overwrites only
' B:H on the existing row. Columns I and J are never written by this
' function under any circumstances, so a user's fulfillment mark or notes
' survive every refresh.
'
' Expects this workbook to live in the same excel-sync\ folder as
' pull_sold_orders.py and config.ini (shared with the lots.xlsm tooling).

' Silent by design (no MsgBox) so it can run unattended from
' Workbook_Open. addedCount/updatedCount are optional ByRef out-params so
' RefreshSoldLotsButton can report a summary; Workbook_Open just calls
' RefreshSoldLots() and ignores them.
Function RefreshSoldLots(Optional ByRef addedCount As Long, Optional ByRef updatedCount As Long) As Boolean
    addedCount = 0
    updatedCount = 0

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sold Lots")

    Dim jsonText As String
    jsonText = RunPullScript()
    If jsonText = "" Then
        RefreshSoldLots = False
        Exit Function
    End If

    If InStr(1, jsonText, Chr(34) & "success" & Chr(34) & ": true") = 0 Then
        RefreshSoldLots = False
        Exit Function
    End If

    Dim orderBlobs As Collection
    Set orderBlobs = ParseOrdersArray(jsonText)

    ' Build a lookup of existing Order ID (column A) -> row number so
    ' matching doesn't re-scan the sheet for every order.
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim existingRows As Object
    Set existingRows = CreateObject("Scripting.Dictionary")
    Dim r As Long
    If lastRow >= 2 Then
        For r = 2 To lastRow
            Dim existingId As String
            existingId = Trim(ws.Cells(r, 1).Value)
            If existingId <> "" And Not existingRows.Exists(existingId) Then
                existingRows.Add existingId, r
            End If
        Next r
    End If

    Dim nextRow As Long
    nextRow = lastRow + 1

    Dim blob As Variant
    For Each blob In orderBlobs
        Dim orderId As String
        orderId = ExtractJSONStringValue(CStr(blob), "id")
        If orderId <> "" Then
            Dim lotId As String, lotName As String, category As String
            Dim buyerEmail As String, statusVal As String
            Dim createdAt As String, paidAt As String
            Dim amountCents As Variant

            lotId = ExtractJSONStringValue(CStr(blob), "lot_id")
            lotName = ExtractJSONStringValue(CStr(blob), "lot_name")
            category = ExtractJSONStringValue(CStr(blob), "category")
            amountCents = ExtractJSONNumberValue(CStr(blob), "amount_cents")
            buyerEmail = ExtractJSONStringValue(CStr(blob), "buyer_email")
            statusVal = ExtractJSONStringValue(CStr(blob), "status")
            createdAt = ExtractJSONStringValue(CStr(blob), "created_at")
            paidAt = ExtractJSONStringValue(CStr(blob), "paid_at")

            Dim salePrice As Variant
            If IsNull(amountCents) Then
                salePrice = ""
            Else
                salePrice = CDbl(amountCents) / 100
            End If

            ' Prefer paid_at (when the sale actually completed); fall back
            ' to created_at for orders that never reached paid status.
            Dim dateSold As String
            dateSold = paidAt
            If dateSold = "" Then dateSold = createdAt

            Dim targetRow As Long
            If existingRows.Exists(orderId) Then
                targetRow = existingRows(orderId)
                updatedCount = updatedCount + 1
            Else
                targetRow = nextRow
                nextRow = nextRow + 1
                existingRows.Add orderId, targetRow
                ws.Cells(targetRow, 1).Value = orderId ' A -- set once on insert, never overwritten
                addedCount = addedCount + 1
            End If

            ' B:H only -- I (Fulfilled) and J (Notes) are local-only and
            ' must never be touched here, on new rows or existing ones.
            ws.Cells(targetRow, 2).Value = dateSold      ' B Date Sold
            ws.Cells(targetRow, 3).Value = lotId          ' C Lot #/SKU
            ws.Cells(targetRow, 4).Value = lotName        ' D Item Name
            ws.Cells(targetRow, 5).Value = category       ' E Category
            ws.Cells(targetRow, 6).Value = salePrice      ' F Sale Price
            ws.Cells(targetRow, 7).Value = buyerEmail     ' G Buyer Email
            ws.Cells(targetRow, 8).Value = statusVal      ' H Order Status
        End If
    Next blob

    RefreshSoldLots = True
End Function

' Manual "Refresh" button -- same engine as the silent Workbook_Open
' refresh, but reports what happened via MsgBox since this one was
' triggered on purpose.
Sub RefreshSoldLotsButton()
    Application.Cursor = xlWait
    Dim added As Long, updated As Long
    Dim success As Boolean
    success = RefreshSoldLots(added, updated)
    Application.Cursor = xlDefault

    If success Then
        MsgBox "Sold Lots refreshed: " & added & " added, " & updated & " updated.", _
            vbInformation, "Refresh Sold Lots"
    Else
        MsgBox "Could not refresh Sold Lots -- check that Python is installed and config.ini is set up correctly.", _
            vbExclamation, "Refresh Sold Lots"
    End If
End Sub

' --- Helpers ----------------------------------------------------------------

Private Function RunPullScript() As String
    Dim shellObj As Object
    Set shellObj = CreateObject("WScript.Shell")

    Dim cmd As String
    cmd = "cmd.exe /c cd /d " & Chr(34) & ThisWorkbook.Path & Chr(34) & _
        " && python pull_sold_orders.py"

    Dim exitCode As Long
    exitCode = shellObj.Run(cmd, 0, True) ' 0 = hidden window, True = wait

    Dim resultPath As String
    resultPath = ThisWorkbook.Path & "\sold_orders.json"

    If Dir(resultPath) = "" Then
        RunPullScript = ""
        Exit Function
    End If

    RunPullScript = ReadTextFile(resultPath)
End Function

' Splits sold_orders.json's "orders": [ {...}, {...}, ... ] array into a
' Collection of raw object text blobs, one per order. Written against the
' exact, predictable shape json.dump(..., indent=2) produces on the
' Python side -- not a general-purpose JSON parser (same house style as
' PublishMacro.ParseRemoteStatusMap in lots.xlsm).
Private Function ParseOrdersArray(ByVal jsonText As String) As Collection
    Dim result As New Collection

    Dim ordersKeyPos As Long
    ordersKeyPos = InStr(1, jsonText, Chr(34) & "orders" & Chr(34))
    If ordersKeyPos = 0 Then
        Set ParseOrdersArray = result
        Exit Function
    End If

    Dim arrayStart As Long
    arrayStart = InStr(ordersKeyPos, jsonText, "[")
    If arrayStart = 0 Then
        Set ParseOrdersArray = result
        Exit Function
    End If

    Dim arrayEnd As Long
    arrayEnd = FindMatchingDelimiter(jsonText, arrayStart, "[", "]")
    If arrayEnd = 0 Then
        Set ParseOrdersArray = result
        Exit Function
    End If

    Dim pos As Long
    pos = arrayStart + 1

    Do While pos < arrayEnd
        Dim objStart As Long
        objStart = InStr(pos, jsonText, "{")
        If objStart = 0 Or objStart > arrayEnd Then Exit Do

        Dim objEnd As Long
        objEnd = FindMatchingDelimiter(jsonText, objStart, "{", "}")
        If objEnd = 0 Then Exit Do

        result.Add Mid(jsonText, objStart, objEnd - objStart + 1)
        pos = objEnd + 1
    Loop

    Set ParseOrdersArray = result
End Function

Private Function FindMatchingDelimiter(ByVal text As String, ByVal openPos As Long, ByVal openChar As String, ByVal closeChar As String) As Long
    Dim depth As Long
    depth = 0
    Dim i As Long
    For i = openPos To Len(text)
        Dim ch As String
        ch = Mid(text, i, 1)
        If ch = openChar Then
            depth = depth + 1
        ElseIf ch = closeChar Then
            depth = depth - 1
            If depth = 0 Then
                FindMatchingDelimiter = i
                Exit Function
            End If
        End If
    Next i
    FindMatchingDelimiter = 0
End Function

' Extracts the string value for "key": "..." out of a JSON object text
' blob, respecting \" and \\ escapes. Returns "" for both a JSON null and
' a genuinely missing key -- fine here since every field this module reads
' is either a string or (for amount_cents) read via ExtractJSONNumberValue
' instead.
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
        ExtractJSONStringValue = "" ' null or non-string -- treat as blank
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

' Extracts a bare numeric value for "key": 1234 (no quotes). Returns Null
' for a JSON null, a missing key, or anything non-numeric.
Private Function ExtractJSONNumberValue(ByVal jsonText As String, ByVal key As String) As Variant
    Dim searchKey As String
    searchKey = Chr(34) & key & Chr(34) & ":"

    Dim startPos As Long
    startPos = InStr(1, jsonText, searchKey)
    If startPos = 0 Then
        ExtractJSONNumberValue = Null
        Exit Function
    End If

    Dim pos As Long
    pos = startPos + Len(searchKey)

    Do While pos <= Len(jsonText) And (Mid(jsonText, pos, 1) = " " Or Mid(jsonText, pos, 1) = vbTab)
        pos = pos + 1
    Loop

    Dim endPos As Long
    endPos = pos
    Do While endPos <= Len(jsonText)
        Dim ch As String
        ch = Mid(jsonText, endPos, 1)
        If ch = "," Or ch = "}" Or ch = vbCr Or ch = vbLf Then Exit Do
        endPos = endPos + 1
    Loop

    Dim rawVal As String
    rawVal = Trim(Mid(jsonText, pos, endPos - pos))

    If rawVal = "" Or LCase(rawVal) = "null" Then
        ExtractJSONNumberValue = Null
    ElseIf IsNumeric(rawVal) Then
        ExtractJSONNumberValue = CDbl(rawVal)
    Else
        ExtractJSONNumberValue = Null
    End If
End Function

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
