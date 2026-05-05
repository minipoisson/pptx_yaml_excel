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
' ThisWorkbook module is used for workbook-level events, such as opening the workbook.

Private Sub Workbook_Open()
  LM_InitializeOnOpen
  ProtectMainSheetUiOnly
End Sub

Private Sub ProtectMainSheetUiOnly()
  On Error GoTo ErrHandler

  With Sheet1
    .Unprotect
    .Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    .EnableSelection = xlUnlockedCells
  End With

  Exit Sub

ErrHandler:
  MsgBox LM_GetText("msg_main_protect_apply_failed"), vbExclamation
End Sub
