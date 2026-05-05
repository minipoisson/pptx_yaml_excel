VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' Main sheet works as the app UI.
' Address-based contract is intentional for transparency and easy inspection:
' B5=source PPT, B8=exported YAML, B11=edited YAML, B14=base PPT,
' B17=output PPT, B20=output folder.
' UI labels are localized by sheet formulas; VBA uses LM_GetText only for dialogs.
'----------------------------------------------------------------


Sub btnSpecifyOrgPpt_Click()
    Dim ppApp As Object
    Dim ppt As Object
    Dim yaml As String
    Dim filePath As String
    Dim yamlPath As String

    filePath = PickFilePath(LM_GetText("dlg_select_ppt"), "PowerPoint Files", "*.ppt; *.pptx")
    If filePath = "" Then
        MsgBox LM_GetText("msg_file_not_selected")
        Exit Sub
    End If
    
    Range("B5").value = filePath
    
    yamlPath = Left(filePath, InStrRev(filePath, ".")) & "yaml"
    Range("B8").value = yamlPath

    Set ppApp = CreateObject("PowerPoint.Application")
    Set ppt = ppApp.Presentations.Open(filePath, , , False)
    
    yaml = ppt2yaml(ppt)
    
    ppt.Close
    ppApp.Quit

    SaveTextAsUtf8NoBom yaml, yamlPath

    MsgBox LM_Format(LM_GetText("msg_yaml_saved"), yamlPath)
End Sub

Sub btnSpecifyYaml_Click()
    Dim yamlPath As String

    yamlPath = PickFilePath(LM_GetText("dlg_select_yaml"), "YAML Files", "*.yaml")
    If yamlPath = "" Then
        MsgBox LM_GetText("msg_file_not_selected")
        Exit Sub
    End If

    Range("B11").value = yamlPath
    UpdateOutputPptPathFromCells
End Sub

Sub btnSpecifyBasePpt_Click()
    Dim basePptPath As String

    basePptPath = PickFilePath(LM_GetText("dlg_select_base_ppt"), "PowerPoint Files", "*.ppt; *.pptx")
    If basePptPath = "" Then
        MsgBox LM_GetText("msg_file_not_selected")
        Exit Sub
    End If

    Range("B14").value = basePptPath
    UpdateOutputPptPathFromCells
End Sub

Sub btnCopyOrgToBase_Click()
    Dim orgPptPath As String

    orgPptPath = Trim(CStr(Range("B5").value))
    If orgPptPath = "" Then
        MsgBox LM_GetText("msg_org_ppt_empty")
        Exit Sub
    End If

    Range("B14").value = orgPptPath
    UpdateOutputPptPathFromCells
End Sub

Sub btnClearAll_Click()
    Range("B5").value = ""
    Range("B8").value = ""
    Range("B11").value = ""
    Range("B14").value = ""
    Range("B17").value = ""
    Range("B20").value = ""
End Sub

Sub btnApplyYamlToPpt_Click()
    On Error GoTo ErrHandler

    Const yamlCellAddress As String = "B11"
    Const baseCellAddress As String = "B14"

    Dim yamlPath As String
    Dim basePptPath As String
    Dim outPptPath As String
    Dim outFolder As String
    Dim yamlText As String
    Dim malformedReport As String
    Dim yamlDict As Object
    Dim ppApp As Object
    Dim ppt As Object
    Dim appliedCount As Long
    Dim pptOnlyCount As Long
    Dim yamlOnlyCount As Long
    Dim pptOnlyReport As String
    Dim yamlOnlyReport As String
    Dim resultMsg As String

    yamlPath = Trim(CStr(Range(yamlCellAddress).value))
    basePptPath = Trim(CStr(Range(baseCellAddress).value))

    If Not ValidateExistingFile(yamlPath, Array("yaml"), LM_Format(LM_GetText("field_modified_yaml"), yamlCellAddress), yamlCellAddress) Then Exit Sub
    If Not ValidateExistingFile(basePptPath, Array("ppt", "pptx"), LM_Format(LM_GetText("field_base_ppt"), baseCellAddress), baseCellAddress) Then Exit Sub

    outPptPath = BuildOutputPptPath(yamlPath, basePptPath)
    Range("B17").value = outPptPath

    If StrComp(basePptPath, outPptPath, vbTextCompare) = 0 Then
        MsgBox LM_GetText("msg_same_output_as_base")
        Exit Sub
    End If

    If Dir(outPptPath) <> "" Then
        If MsgBox(LM_Format(LM_GetText("msg_confirm_overwrite"), outPptPath), vbYesNo + vbDefaultButton2) <> vbYes Then
            Exit Sub
        End If
        Kill outPptPath
    End If

    FileCopy basePptPath, outPptPath

    yamlText = ReadTextUtf8(yamlPath)
    Set yamlDict = ParseYamlToDictionary(yamlText, malformedReport)

    Set ppApp = CreateObject("PowerPoint.Application")
    Set ppt = ppApp.Presentations.Open(outPptPath, , , False)
    ApplyYamlToPresentation ppt, yamlDict, appliedCount, pptOnlyCount, yamlOnlyCount, pptOnlyReport, yamlOnlyReport
    ppt.Save
    ppt.Close
    ppApp.Quit
    Set ppt = Nothing
    Set ppApp = Nothing

    outFolder = Left(outPptPath, InStrRev(outPptPath, "\") - 1)
    Range("B20").value = outFolder
    OpenOutputLocation outPptPath, outFolder

    resultMsg = LM_GetText("msg_apply_completed") & vbCrLf & _
                LM_Format(LM_GetText("msg_applied_count"), appliedCount) & vbCrLf & _
                LM_Format(LM_GetText("msg_ppt_only_count"), pptOnlyCount) & vbCrLf & _
                LM_Format(LM_GetText("msg_yaml_only_count"), yamlOnlyCount)

    If Len(malformedReport) > 0 Then
        resultMsg = resultMsg & vbCrLf & vbCrLf & LM_GetText("msg_malformed_yaml") & vbCrLf & FirstNLines(malformedReport, 20)
    End If
    If Len(pptOnlyReport) > 0 Then
        resultMsg = resultMsg & vbCrLf & vbCrLf & LM_GetText("msg_ppt_only_keys") & vbCrLf & pptOnlyReport
    End If
    If Len(yamlOnlyReport) > 0 Then
        resultMsg = resultMsg & vbCrLf & vbCrLf & LM_GetText("msg_yaml_only_keys") & vbCrLf & yamlOnlyReport
    End If

    MsgBox resultMsg
    Exit Sub

ErrHandler:
    On Error Resume Next
    If Not ppt Is Nothing Then ppt.Close
    If Not ppApp Is Nothing Then ppApp.Quit
    On Error GoTo 0
    MsgBox LM_Format(LM_GetText("msg_apply_error"), Err.Description)
End Sub

Private Function PickFilePath(dialogTitle As String, filterLabel As String, filterPattern As String) As String
    Dim fDialog As FileDialog

    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    With fDialog
        .Title = dialogTitle
        .InitialFileName = ThisWorkbook.Path & "\"
        .Filters.Clear
        .Filters.Add filterLabel, filterPattern, 1
        If .Show = -1 Then
            PickFilePath = .SelectedItems(1)
        Else
            PickFilePath = ""
        End If
    End With
End Function

Private Sub UpdateOutputPptPathFromCells()
    Dim yamlPath As String
    Dim basePptPath As String

    yamlPath = Trim(CStr(Range("B11").value))
    basePptPath = Trim(CStr(Range("B14").value))

    If yamlPath = "" Or basePptPath = "" Then
        Range("B17").value = ""
        Exit Sub
    End If

    If InStrRev(yamlPath, ".") = 0 Or InStrRev(basePptPath, ".") = 0 Then
        Range("B17").value = ""
        Exit Sub
    End If

    Range("B17").value = BuildOutputPptPath(yamlPath, basePptPath)
End Sub

Private Function ValidateExistingFile(filePath As String, allowedExts As Variant, fieldLabel As String, cellAddress As String) As Boolean
    Dim ext As String
    Dim i As Long
    Dim okExt As Boolean

    ValidateExistingFile = False
    If filePath = "" Then
        MsgBox LM_Format(LM_GetText("msg_field_empty"), fieldLabel)
        Range(cellAddress).Select
        Exit Function
    End If

    If Dir(filePath) = "" Then
        MsgBox LM_Format(LM_GetText("msg_file_not_found"), fieldLabel, filePath)
        Range(cellAddress).Select
        Exit Function
    End If

    If InStrRev(filePath, ".") = 0 Then
        MsgBox LM_Format(LM_GetText("msg_invalid_extension"), fieldLabel, filePath)
        Range(cellAddress).Select
        Exit Function
    End If

    ext = LCase(Mid(filePath, InStrRev(filePath, ".") + 1))
    okExt = False
    For i = LBound(allowedExts) To UBound(allowedExts)
        If ext = LCase(CStr(allowedExts(i))) Then
            okExt = True
            Exit For
        End If
    Next i

    If Not okExt Then
        MsgBox LM_Format(LM_GetText("msg_invalid_extension"), fieldLabel, filePath)
        Range(cellAddress).Select
        Exit Function
    End If

    ValidateExistingFile = True
End Function

Private Sub OpenOutputLocation(outPptPath As String, outFolder As String)
    On Error Resume Next
    Shell "explorer.exe /select,""" & outPptPath & """", vbNormalFocus
    If Err.Number <> 0 Then
        Err.Clear
        Shell "explorer.exe """ & outFolder & """", vbNormalFocus
    End If
    On Error GoTo 0
End Sub

Private Function FirstNLines(textValue As String, maxLines As Long) As String
    Dim lines() As String
    Dim i As Long
    Dim n As Long
    Dim output As String

    If Len(textValue) = 0 Then
        FirstNLines = ""
        Exit Function
    End If

    lines = Split(textValue, vbCrLf)
    n = UBound(lines)
    If n > maxLines - 1 Then n = maxLines - 1

    For i = 0 To n
        If Len(lines(i)) > 0 Then
            output = output & lines(i) & vbCrLf
        End If
    Next i

    FirstNLines = output
End Function




