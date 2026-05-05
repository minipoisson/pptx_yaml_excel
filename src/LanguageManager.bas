Attribute VB_Name = "LanguageManager"
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function GetLocaleInfoA Lib "kernel32" ( _
    ByVal Locale As Long, _
    ByVal LCType As Long, _
    ByVal lpLCData As String, _
    ByVal cchData As Long) As Long
#Else
Private Declare Function GetLocaleInfoA Lib "kernel32" ( _
    ByVal Locale As Long, _
    ByVal LCType As Long, _
    ByVal lpLCData As String, _
    ByVal cchData As Long) As Long
#End If

Private Const LOCALE_SNAME As Long = &H5C

' -----------------------------------------------------------------
' LM_InitializeOnOpen
'   Call from Workbook_Open.
'   Detects the system language and sets Sheet1!D1 to the
'   matching display name (e.g. "English" / "日本語").
'   Defaults to "English" if detection fails.
' -----------------------------------------------------------------
Public Sub LM_InitializeOnOpen()
    On Error GoTo ErrDefault

    Sheet1.Range("D1").value = LM_DetectSystemLanguageName()
    Exit Sub

ErrDefault:
    On Error Resume Next
    Sheet1.Range("D1").value = "English"
End Sub

' -----------------------------------------------------------------
' LM_GetText
'   Looks up keyId in the i18n sheet using the column index from
'   Sheet1!F1 (=MATCH(D1, i18n row 2, 0)) for the current language.
'   Falls back to the English column (col 2), then to keyId itself.
'   Converts literal "\n" in stored values to vbCrLf.
' -----------------------------------------------------------------
Public Function LM_GetText(ByVal keyId As String) As String
    On Error GoTo Fallback

    ' Sheet1!F1 is maintained by worksheet formulas from the language name in D1.
    ' VBA only consumes that resolved column index.
    Dim langCol As Long
    langCol = CLng(Sheet1.Range("F1").value)
    If langCol < 2 Then langCol = 2

    Dim ws As Worksheet
    Set ws = LM_I18nSheet()
    If ws Is Nothing Then GoTo Fallback

    Dim foundRow As Long
    foundRow = LM_FindKeyRow(ws, keyId)
    If foundRow = 0 Then GoTo Fallback

    Dim result As String
    result = Trim$(CStr(ws.Cells(foundRow, langCol).value))

    If Len(result) = 0 Then
        result = Trim$(CStr(ws.Cells(foundRow, 2).value))
    End If

    If Len(result) > 0 Then
        LM_GetText = Replace(result, "\n", vbCrLf)
        Exit Function
    End If

Fallback:
    LM_GetText = keyId
End Function

' -----------------------------------------------------------------
' LM_Format
'   Replaces {0}, {1}, ... placeholders with supplied arguments.
' -----------------------------------------------------------------
Public Function LM_Format(ByVal templateValue As String, ParamArray args() As Variant) As String
    Dim output As String
    Dim i As Long
    output = templateValue
    On Error GoTo Done
    For i = LBound(args) To UBound(args)
        output = Replace(output, "{" & CStr(i) & "}", CStr(args(i)))
    Next i
Done:
    LM_Format = output
End Function

' -----------------------------------------------------------------
' Private helpers
' -----------------------------------------------------------------

Private Function LM_DetectSystemLanguageName() As String
    ' 1. Get LCID from Office.
    Dim lcid As Long
    On Error Resume Next
    lcid = Application.LanguageSettings.LanguageID(2)
    On Error GoTo 0

    ' 2. Get BCP 47 locale name via Windows API (e.g. "ja-JP", "zh-CN", "zh-TW").
    Dim localeName As String
    If lcid > 0 Then
        Dim buf As String
        buf = Space$(85)           ' LOCALE_NAME_MAX_LENGTH = 85
        Dim ret As Long
        ret = GetLocaleInfoA(lcid, LOCALE_SNAME, buf, Len(buf))
        If ret > 1 Then
            localeName = LCase$(Left$(buf, ret - 1))  ' ret includes null terminator
        End If
    End If

    Dim ws As Worksheet
    Set ws = LM_I18nSheet()

    ' 3. Match against i18n row 1 in two passes:
    '    Pass 1 ? exact match        ("zh-cn" matches i18n code "zh-cn")
    '    Pass 2 ? language-prefix    ("ja-jp" matches i18n code "ja")
    If Len(localeName) > 0 Then
        Dim lastCol As Long
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        Dim col As Long
        Dim cellCode As String

        ' Pass 1: exact
        For col = 2 To lastCol
            cellCode = LCase$(Trim$(CStr(ws.Cells(1, col).value)))
            If cellCode = localeName Then
                LM_DetectSystemLanguageName = CStr(ws.Cells(2, col).value)
                Exit Function
            End If
        Next col

        ' Pass 2: i18n code is a prefix of the locale name
        '         e.g. localeName="ja-jp", cellCode="ja"
        For col = 2 To lastCol
            cellCode = LCase$(Trim$(CStr(ws.Cells(1, col).value)))
            If Len(cellCode) > 0 Then
                If Left$(localeName, Len(cellCode) + 1) = cellCode & "-" Then
                    LM_DetectSystemLanguageName = CStr(ws.Cells(2, col).value)
                    Exit Function
                End If
            End If
        Next col
    End If

    ' 4. Language not listed in i18n ? fall back to default.
    LM_DetectSystemLanguageName = LM_DefaultLanguageName(ws)
End Function

Private Function LM_DefaultLanguageName(ByVal ws As Worksheet) As String
    ' Returns the row-2 display name for the default language code ("en").
    ' "en" is the only hardcoded value; the display name comes from the i18n sheet.
    ' Falls back to the literal "English" only when the i18n sheet is unavailable.
    Const DEFAULT_LANG_CODE As String = "en"
    Const DEFAULT_LANG_NAME As String = "English"  ' last-resort hardcoded name

    If Not ws Is Nothing Then
        Dim lastCol As Long
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        Dim col As Long
        For col = 2 To lastCol
            If LCase$(Trim$(CStr(ws.Cells(1, col).value))) = DEFAULT_LANG_CODE Then
                LM_DefaultLanguageName = CStr(ws.Cells(2, col).value)
                Exit Function
            End If
        Next col
    End If

    LM_DefaultLanguageName = DEFAULT_LANG_NAME
End Function

Private Function LM_I18nSheet() As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, "i18n", vbTextCompare) = 0 Then
            Set LM_I18nSheet = ws
            Exit Function
        End If
    Next ws
End Function

Private Function LM_FindKeyRow(ByVal ws As Worksheet, ByVal keyId As String) As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 3 Then Exit Function
    Dim r As Long
    For r = 3 To lastRow
        If StrComp(CStr(ws.Cells(r, 1).value), keyId, vbTextCompare) = 0 Then
            LM_FindKeyRow = r
            Exit Function
        End If
    Next r
End Function


