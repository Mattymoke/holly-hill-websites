VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' Two-way auto-sync on save: whenever this workbook is saved, every changed
' row gets pushed to the website, and every unchanged row gets checked
' against the website in case its status changed there (e.g. a lot sold
' through the shop, not through Excel). Uses PublishRow and
' PullRemoteStatus (both in the PublishMacro module) as the shared push/pull
' engines.
'
' Save flow:
'   Phase 1 (pre-pass) -- pull every lot's current status from the website
'     once, before touching any row. If the pull fails for any reason, skip
'     reconciliation entirely for this save and fall back to push-only
'     behavior, same as before this feature existed.
'   Phase 2 (push) -- for each row whose fingerprint changed: if the
'     website already shows it "sold" but the local Status cell doesn't,
'     correct the local cell to "sold" before building the payload (so we
'     push the corrected value, not overwrite the website's sold status
'     with a stale local one). Otherwise push normally.
'   Phase 3 (reconcile) -- for rows that were NOT pushed this save (nothing
'     changed locally): if the website's status differs from the local
'     Status cell, pull it down into the sheet and refresh Last Synced /
'     the fingerprint so it isn't flagged as a pending local change next
'     time.
'
' Columns (Lots sheet):
'   A-K  Lot #, Category, Item Name, Description, Price, Status, Photo 1-5
'   L    Last Synced (timestamp)
'   M    Sync Status ("OK" / "ERROR: ..." / a reconciliation note)
'   N    Last Synced Fingerprint (hidden)
'   O    Last Synced Photo Paths (hidden)

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    On Error GoTo Finish

    Dim ws As Worksheet
    Set ws = Me.Worksheets("Lots")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then GoTo Finish ' no data rows yet

    ' --- Phase 1: pull remote status once, before touching any row --------

    Dim remoteStatus As Object
    Set remoteStatus = PullRemoteStatus() ' Nothing if the pull failed

    Dim pulledOk As Boolean
    pulledOk = Not (remoteStatus Is Nothing)

    Dim syncedCount As Long, pulledCount As Long, failedCount As Long
    syncedCount = 0
    pulledCount = 0
    failedCount = 0

    Dim r As Long
    For r = 2 To lastRow
        If Trim(ws.Cells(r, 1).Value) <> "" Then ' skip genuinely blank rows

            Dim lotId As String
            lotId = Trim(ws.Cells(r, 1).Value)

            Dim currentFingerprint As String
            currentFingerprint = BuildFingerprint(ws, r)

            Dim storedFingerprint As String
            storedFingerprint = CStr(ws.Cells(r, 14).Value) ' N

            Dim syncStatusVal As String
            syncStatusVal = CStr(ws.Cells(r, 13).Value) ' M

            Dim needsSync As Boolean
            needsSync = True
            If currentFingerprint = storedFingerprint Then
                ' Nothing changed since last sync -- only re-run if the
                ' last attempt actually failed.
                If syncStatusVal <> "" And Left(syncStatusVal, 5) <> "ERROR" Then
                    needsSync = False
                End If
            End If

            If needsSync Then
                ' --- Phase 2: push, with the sold-status guard -----------

                Dim statusWasCorrected As Boolean
                statusWasCorrected = False

                If pulledOk Then
                    If remoteStatus.Exists(lotId) Then
                        If remoteStatus(lotId) = "sold" And LCase(Trim(ws.Cells(r, 6).Value)) <> "sold" Then
                            ws.Cells(r, 6).Value = "sold"
                            statusWasCorrected = True
                            currentFingerprint = BuildFingerprint(ws, r) ' reflect the corrected status
                        End If
                    End If
                End If

                Dim currentPathsCSV As String
                currentPathsCSV = CombinePhotoColumns(ws, r)

                Dim previousPathsCSV As String
                previousPathsCSV = Trim(ws.Cells(r, 15).Value) ' O

                Dim clearImages As Boolean
                Dim pathsToSend As String

                If currentPathsCSV = "" And previousPathsCSV <> "" Then
                    ' All photos removed from the sheet -- tell the server
                    ' to clear them too, instead of leaving stale images.
                    clearImages = True
                    pathsToSend = ""
                Else
                    clearImages = False
                    pathsToSend = GetNewPaths(currentPathsCSV, previousPathsCSV)
                End If

                Dim success As Boolean
                success = PublishRow(ws, r, pathsToSend, clearImages)

                If success Then
                    ws.Cells(r, 14).Value = currentFingerprint
                    ws.Cells(r, 15).Value = currentPathsCSV
                    If statusWasCorrected Then
                        ws.Cells(r, 13).Value = "Website shows this sold -- local status corrected before syncing."
                    End If
                    syncedCount = syncedCount + 1
                Else
                    failedCount = failedCount + 1
                End If
            Else
                ' --- Phase 3: reconcile rows that were NOT pushed ---------

                If pulledOk Then
                    If remoteStatus.Exists(lotId) Then
                        Dim remoteVal As String
                        remoteVal = remoteStatus(lotId)
                        If remoteVal <> "" And LCase(remoteVal) <> LCase(Trim(ws.Cells(r, 6).Value)) Then
                            ws.Cells(r, 6).Value = remoteVal
                            ws.Cells(r, 12).Value = Now
                            ws.Cells(r, 14).Value = BuildFingerprint(ws, r)
                            ws.Cells(r, 13).Value = "Updated from website: " & remoteVal
                            pulledCount = pulledCount + 1
                        End If
                    End If
                End If
            End If
        End If
    Next r

    If syncedCount > 0 Or pulledCount > 0 Or failedCount > 0 Then
        Dim parts As New Collection
        If syncedCount > 0 Then parts.Add "Synced " & syncedCount & " row(s)"
        If pulledCount > 0 Then parts.Add "Updated " & pulledCount & " from website"
        If failedCount > 0 Then parts.Add failedCount & " failed"

        Dim msg As String
        msg = ""
        Dim p As Variant
        Dim idx As Long
        idx = 0
        For Each p In parts
            If idx > 0 Then msg = msg & ", "
            msg = msg & p
            idx = idx + 1
        Next p

        If Not pulledOk Then
            msg = msg & vbCrLf & "(Could not check the website for status updates this save -- push-only.)"
        End If

        MsgBox msg, vbInformation, "Auto-Sync"
    End If

Finish:
    Cancel = False ' the file save must always proceed, sync problems or not
End Sub
