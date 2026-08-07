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

' Auto-refreshes the Sold Lots sheet every time this workbook is opened,
' silently -- no MsgBox, so opening the file just shows current data
' without interrupting anyone. Uses RefreshSoldLots (in SoldLotsMacro) as
' the shared pull engine. This workbook is pull-only: nothing it does
' ever writes back to the website.
'
' If the pull fails for any reason (no network, Python missing, bad
' config), Workbook_Open fails silently too -- the sheet just shows
' whatever was there from the last successful refresh, which is a much
' better failure mode than an error dialog on every open.

Private Sub Workbook_Open()
    On Error Resume Next
    RefreshSoldLots
    On Error GoTo 0
End Sub
