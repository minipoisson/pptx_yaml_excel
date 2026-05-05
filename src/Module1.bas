Attribute VB_Name = "Module1"
Option Explicit

Private Const SHAPE_TYPE_GROUP As Long = 6
Private Const SHAPE_TYPE_TABLE As Long = 19

Function ppt2yaml(ppt As Object) As String
    ' YAML keys are intentionally derived from slide structure and shape names.
    ' This is suitable for translation workflows, but not a stable diff format for
    ' iterative deck editing; re-export after structural edits.
    Dim yaml As String
    Dim sld As Object
    Dim para As Object
    Dim i As Integer
    Dim notesShape As Object

    yaml = ""

    For Each sld In ppt.Slides
        ExportShapesRecursive sld.Shapes, sld.slideIndex, "", yaml

        Set notesShape = Nothing
        On Error Resume Next
        ' NotesPage.Shapes(2) is the standard notes placeholder in typical PPT layouts.
        ' If unavailable, notes are skipped and processing continues.
        Set notesShape = sld.NotesPage.Shapes(2)
        On Error GoTo 0

        If Not notesShape Is Nothing Then
            If notesShape.HasTextFrame Then
                If notesShape.TextFrame.HasText Then
                    For i = 1 To notesShape.TextFrame.TextRange.Paragraphs.Count
                        Set para = notesShape.TextFrame.TextRange.Paragraphs(i)
                        AppendParagraphYaml para, sld.slideIndex, "notes", i - 1, yaml
                    Next i
                End If
            End If
        End If
    Next sld

    ppt2yaml = yaml
End Function

Private Sub ExportShapesRecursive(ByVal shpCollection As Object, ByVal slideIndex As Long, ByVal parentPath As String, ByRef yaml As String)
    Dim shp As Object
    Dim para As Object
    Dim i As Integer
    Dim pathSegment As String
    Dim currentPath As String
    Dim siblingCounter As Object
    Dim tblExport As Object
    Dim rExport As Long
    Dim cExport As Long
    Dim jExport As Long
    Dim cellPathExport As String

    Set siblingCounter = CreateObject("Scripting.Dictionary")

    For Each shp In shpCollection
        pathSegment = BuildUniqueShapeSegment(shp, siblingCounter)

        If Len(parentPath) > 0 Then
            currentPath = parentPath & "." & pathSegment
        Else
            currentPath = pathSegment
        End If

        If shp.Type = SHAPE_TYPE_GROUP Then
            If shp.GroupItems.Count > 0 Then
                ExportShapesRecursive shp.GroupItems, slideIndex, currentPath, yaml
            End If
        ElseIf shp.Type = SHAPE_TYPE_TABLE Then
            Set tblExport = shp.Table
            For rExport = 0 To tblExport.Rows.Count - 1
                For cExport = 0 To tblExport.Columns.Count - 1
                    With tblExport.Cell(rExport + 1, cExport + 1).Shape
                        If .HasTextFrame Then
                            If .TextFrame.HasText Then
                                cellPathExport = currentPath & ".r" & Format(rExport, "00") & "c" & Format(cExport, "00")
                                For jExport = 1 To .TextFrame.TextRange.Paragraphs.Count
                                    AppendParagraphYaml .TextFrame.TextRange.Paragraphs(jExport), slideIndex, cellPathExport, jExport - 1, yaml
                                Next jExport
                            End If
                        End If
                    End With
                Next cExport
            Next rExport
        ElseIf shp.HasTextFrame Then
            If shp.TextFrame.HasText Then
                For i = 1 To shp.TextFrame.TextRange.Paragraphs.Count
                    Set para = shp.TextFrame.TextRange.Paragraphs(i)
                    AppendParagraphYaml para, slideIndex, currentPath, i - 1, yaml
                Next i
            End If
        End If
    Next shp

End Sub

Private Sub AppendParagraphYaml(ByVal para As Object, ByVal slideIndex As Long, ByVal shapePath As String, ByVal paraIndex As Long, ByRef yaml As String)
    Dim value As String
    Dim key As String

    value = CStr(para.Text)

    If Len(Trim(value)) = 0 Then Exit Sub

    value = EscapeYamlValue(value)
    key = BuildYamlKey(slideIndex, shapePath, paraIndex)

    yaml = yaml & key & ": """ & value & """" & vbCrLf
End Sub

Private Function EscapeYamlValue(ByVal value As String) As String
    value = Replace(value, "\", "\\")
    value = Replace(value, """", "\""")
    value = Replace(value, Chr(11), "\n")
    value = Replace(value, vbCr, "\n")

    Do While Right(value, 2) = "\n"
        value = Left(value, Len(value) - 2)
    Loop

    EscapeYamlValue = value
End Function

Private Function BuildYamlKey(ByVal slideIndex As Long, ByVal shapePath As String, ByVal paraIndex As Long) As String
    BuildYamlKey = "s" & Format(slideIndex, "00") & "." & shapePath & ".p" & Format(paraIndex, "00")
End Function

Private Function BuildUniqueShapeSegment(ByVal shp As Object, ByVal siblingCounter As Object) As String
    Dim baseName As String
    Dim currentCount As Long

    ' Duplicate shape names are disambiguated by sibling order (_2, _3, ...).
    baseName = NormalizeShapeNamePart(CStr(shp.Name))
    If Len(baseName) = 0 Then baseName = "shape"

    If siblingCounter.Exists(baseName) Then
        currentCount = CLng(siblingCounter(baseName)) + 1
        siblingCounter(baseName) = currentCount
    Else
        currentCount = 1
        siblingCounter(baseName) = currentCount
    End If

    If currentCount = 1 Then
        BuildUniqueShapeSegment = baseName
    Else
        BuildUniqueShapeSegment = baseName & "_" & CStr(currentCount)
    End If
End Function

Private Function NormalizeShapeNamePart(ByVal namePart As String) As String
    NormalizeShapeNamePart = Trim(namePart)
    NormalizeShapeNamePart = Replace(NormalizeShapeNamePart, " ", "_")
    NormalizeShapeNamePart = Replace(NormalizeShapeNamePart, ".", "_")
End Function



Public Sub SaveTextAsUtf8NoBom(content As String, filePath As String)
    Dim st As Object
    Dim stBin As Object

    Set st = CreateObject("ADODB.Stream")
    st.Type = 2          ' text mode
    st.Charset = "UTF-8"
    st.Open
    st.WriteText content
    st.Position = 0
    st.Type = 1          ' switch to binary mode
    st.Position = 3      ' skip the first 3 bytes (BOM)

    Set stBin = CreateObject("ADODB.Stream")
    stBin.Type = 1
    stBin.Open
    st.CopyTo stBin
    stBin.SaveToFile filePath, 2

    stBin.Close
    st.Close
End Sub

Public Function ReadTextUtf8(filePath As String) As String
    Dim st As Object

    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "UTF-8"
    st.Open
    st.LoadFromFile filePath
    ReadTextUtf8 = st.ReadText
    st.Close
End Function

Public Function BuildOutputPptPath(yamlPath As String, basePptPath As String) As String
    Dim outDir As String
    Dim yamlBaseName As String
    Dim ext As String

    outDir = Left(yamlPath, InStrRev(yamlPath, "\\"))
    yamlBaseName = Mid(yamlPath, InStrRev(yamlPath, "\\") + 1)
    yamlBaseName = Left(yamlBaseName, InStrRev(yamlBaseName, ".") - 1)
    ext = Mid(basePptPath, InStrRev(basePptPath, "."))

    BuildOutputPptPath = outDir & yamlBaseName & ext
End Function

Public Function ParseYamlToDictionary(yamlText As String, ByRef malformedReport As String) As Object
    Dim dict As Object
    Dim lines() As String
    Dim lineText As String
    Dim key As String
    Dim rawValue As String
    Dim valueText As String
    Dim sepPos As Long
    Dim i As Long
    Dim escTemp As String

    Set dict = CreateObject("Scripting.Dictionary")
    escTemp = Chr(1)  ' Use a character that does not appear in PPT text as a placeholder
    malformedReport = ""
    lines = Split(yamlText, vbLf)

    For i = LBound(lines) To UBound(lines)
        lineText = Replace(lines(i), vbCr, "")

        If Len(Trim(lineText)) = 0 Then GoTo lblContinue
        If Left(Trim(lineText), 1) = "#" Then GoTo lblContinue

        sepPos = InStr(lineText, ": ")
        If sepPos <= 1 Then
            malformedReport = malformedReport & "Line " & (i + 1) & ": " & lineText & vbCrLf
            GoTo lblContinue
        End If

        key = Left(lineText, sepPos - 1)
        rawValue = Mid(lineText, sepPos + 2)

        If Len(rawValue) < 2 Or Left(rawValue, 1) <> """" Or Right(rawValue, 1) <> """" Then
            malformedReport = malformedReport & "Line " & (i + 1) & ": " & lineText & vbCrLf
            GoTo lblContinue
        End If

        valueText = Mid(rawValue, 2, Len(rawValue) - 2)
        valueText = Replace(valueText, "\\\\", escTemp)         ' \\ → placeholder (process first)
        valueText = Replace(valueText, "\\" & """", """")    ' \" → "
        valueText = Replace(valueText, "\n", Chr(11))           ' \n → soft line break
        valueText = Replace(valueText, escTemp, "\\")          ' placeholder → \

        If dict.Exists(key) Then
            malformedReport = malformedReport & "Line " & (i + 1) & " (duplicate key): " & key & vbCrLf
        End If
        dict(key) = valueText

lblContinue:
    Next i

    Set ParseYamlToDictionary = dict
End Function

Public Sub ApplyYamlToPresentation(ByVal ppt As Object, ByVal yamlDict As Object, _
                                   ByRef appliedCount As Long, ByRef pptOnlyCount As Long, _
                                   ByRef yamlOnlyCount As Long, ByRef pptOnlyReport As String, _
                                   ByRef yamlOnlyReport As String)
    ' Keys missing on either side are counted and reported, not treated as fatal.
    Dim sld As Object
    Dim para As Object
    Dim i As Integer
    Dim key As String
    Dim matched As Object
    Dim k As Variant
    Dim notesShp As Object

    Set matched = CreateObject("Scripting.Dictionary")
    appliedCount = 0
    pptOnlyCount = 0
    yamlOnlyCount = 0
    pptOnlyReport = ""
    yamlOnlyReport = ""

    For Each sld In ppt.Slides
        ApplyShapesRecursive sld.Shapes, sld.slideIndex, "", yamlDict, matched, appliedCount, pptOnlyCount, pptOnlyReport

        Set notesShp = Nothing
        On Error Resume Next
        ' Keep behavior symmetric with export: notes are optional.
        Set notesShp = sld.NotesPage.Shapes(2)
        On Error GoTo 0

        If Not notesShp Is Nothing Then
            If notesShp.HasTextFrame Then
                If notesShp.TextFrame.HasText Then
                    For i = 1 To notesShp.TextFrame.TextRange.Paragraphs.Count
                        Set para = notesShp.TextFrame.TextRange.Paragraphs(i)
                        key = BuildYamlKey(sld.slideIndex, "notes", i - 1)
                        ApplyYamlToParagraph para, key, yamlDict, matched, appliedCount, pptOnlyCount, pptOnlyReport
                    Next i
                End If
            End If
        End If
    Next sld

    For Each k In yamlDict.Keys
        If Not matched.Exists(CStr(k)) Then
            yamlOnlyCount = yamlOnlyCount + 1
            If yamlOnlyCount <= 20 Then
                yamlOnlyReport = yamlOnlyReport & CStr(k) & vbCrLf
            End If
        End If
    Next k
End Sub

Private Sub ApplyShapesRecursive(ByVal shpCollection As Object, ByVal slideIndex As Long, ByVal parentPath As String, _
                                 ByVal yamlDict As Object, ByVal matched As Object, ByRef appliedCount As Long, _
                                 ByRef pptOnlyCount As Long, ByRef pptOnlyReport As String)
    Dim shp As Object
    Dim para As Object
    Dim i As Integer
    Dim pathSegment As String
    Dim currentPath As String
    Dim key As String
    Dim siblingCounter As Object
    Dim tblApply As Object
    Dim rApply As Long
    Dim cApply As Long
    Dim jApply As Long
    Dim cellPathApply As String
    Dim cellKeyApply As String

    Set siblingCounter = CreateObject("Scripting.Dictionary")

    For Each shp In shpCollection
        pathSegment = BuildUniqueShapeSegment(shp, siblingCounter)

        If Len(parentPath) > 0 Then
            currentPath = parentPath & "." & pathSegment
        Else
            currentPath = pathSegment
        End If

        If shp.Type = SHAPE_TYPE_GROUP Then
            If shp.GroupItems.Count > 0 Then
                ApplyShapesRecursive shp.GroupItems, slideIndex, currentPath, yamlDict, matched, appliedCount, pptOnlyCount, pptOnlyReport
            End If
        ElseIf shp.Type = SHAPE_TYPE_TABLE Then
            Set tblApply = shp.Table
            For rApply = 0 To tblApply.Rows.Count - 1
                For cApply = 0 To tblApply.Columns.Count - 1
                    With tblApply.Cell(rApply + 1, cApply + 1).Shape
                        If .HasTextFrame Then
                            If .TextFrame.HasText Then
                                cellPathApply = currentPath & ".r" & Format(rApply, "00") & "c" & Format(cApply, "00")
                                For jApply = 1 To .TextFrame.TextRange.Paragraphs.Count
                                    cellKeyApply = BuildYamlKey(slideIndex, cellPathApply, jApply - 1)
                                    ApplyYamlToParagraph .TextFrame.TextRange.Paragraphs(jApply), cellKeyApply, yamlDict, matched, appliedCount, pptOnlyCount, pptOnlyReport
                                Next jApply
                            End If
                        End If
                    End With
                Next cApply
            Next rApply
        ElseIf shp.HasTextFrame Then
            If shp.TextFrame.HasText Then
                For i = 1 To shp.TextFrame.TextRange.Paragraphs.Count
                    Set para = shp.TextFrame.TextRange.Paragraphs(i)
                    key = BuildYamlKey(slideIndex, currentPath, i - 1)
                    ApplyYamlToParagraph para, key, yamlDict, matched, appliedCount, pptOnlyCount, pptOnlyReport
                Next i
            End If
        End If
    Next shp
End Sub

Private Sub ApplyYamlToParagraph(ByVal para As Object, ByVal key As String, ByVal yamlDict As Object, ByVal matched As Object, _
                                 ByRef appliedCount As Long, ByRef pptOnlyCount As Long, ByRef pptOnlyReport As String)
    If Len(Trim(CStr(para.Text))) = 0 Then Exit Sub

    If yamlDict.Exists(key) Then
        para.Text = yamlDict(key)
        matched(key) = True
        appliedCount = appliedCount + 1
    Else
        pptOnlyCount = pptOnlyCount + 1
        If pptOnlyCount <= 20 Then
            pptOnlyReport = pptOnlyReport & key & vbCrLf
        End If
    End If
End Sub


