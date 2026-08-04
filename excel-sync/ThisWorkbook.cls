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

' Auto-sync on save: whenever this workbook is saved, every changed data
' row gets published automatically -- no button required. Uses PublishRow
' (in the PublishMacro module) as the shared engine, and only re-uploads
' photos that weren't already synced last time (tracked via the hidden
' J / K columns below).
'
' Columns (Lots sheet):
'   A-G  Lot #, Category, Item Name, Description, Price, Status, Photo Paths
'   H    Last Synced (timestamp)
'   I    Sync Status ("OK" or "ERROR: ...")
'   J    Last Synced Fingerprint (hidden) -- A-G as of the last successful/attempted sync
'   K    Last Synced Photo Paths (hidden) -- photo paths already uploaded as of last sync

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    On Error GoTo Finish

    Dim ws As Worksheet
    Set ws = Me.Worksheets("Lots")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then GoTo Finish ' no data rows yet

    Dim syncedCount As Long, failedCount As Long
    syncedCount = 0
    failedCount = 0

    Dim r As Long
    For r = 2 To lastRow
        If Trim(ws.Cells(r, 1).Value) <> "" Then ' skip genuinely blank rows

            Dim currentFingerprint As String
            currentFingerprint = BuildFingerprint(ws, r)

            Dim storedFingerprint As String
            storedFingerprint = CStr(ws.Cells(r, 10).Value) ' J

            Dim syncStatusVal As String
            syncStatusVal = CStr(ws.Cells(r, 9).Value) ' I

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
                Dim currentPathsCSV As String
                currentPathsCSV = Trim(ws.Cells(r, 7).Value) ' G

                Dim previousPathsCSV As String
                previousPathsCSV = Trim(ws.Cells(r, 11).Value) ' K

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
                    ws.Cells(r, 10).Value = currentFingerprint
                    ws.Cells(r, 11).Value = currentPathsCSV
                    syncedCount = syncedCount + 1
                Else
                    failedCount = failedCount + 1
                End If
            End If
        End If
    Next r

    If syncedCount > 0 Or failedCount > 0 Then
        Dim msg As String
        msg = "Synced " & syncedCount & " row(s)"
        If failedCount > 0 Then msg = msg & ", " & failedCount & " failed"
        MsgBox msg, vbInformation, "Auto-Sync"
    End If

Finish:
    Cancel = False ' the file save must always proceed, sync problems or not
End Sub
