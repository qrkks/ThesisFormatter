Private Const PAGE_NUMBER_MODE_ARABIC_ONLY As Integer = 1
Private Const PAGE_NUMBER_MODE_MIXED_BY_TOC As Integer = 2

' 处理标题的格式化宏
Sub FormatTitleByHeadingStyle()
    Dim para As Paragraph
    For Each para In ActiveDocument.Paragraphs
        If para.Style = "标题" Then
            With para.Range
                .Font.NameFarEast = "黑体"
                .Font.Name = "黑体"
                .Font.Size = 18 ' 小二
                .Font.Bold = True
                .Font.Color = wdColorBlack
                .ParagraphFormat.Alignment = wdAlignParagraphCenter
            End With
        End If
    Next para
    ' MsgBox "题目格式化完成！"
End Sub

' 一级标题格式化宏
Sub FormatLevel1Heading()
    Dim para As Paragraph
    For Each para In ActiveDocument.Paragraphs
        If para.Style = "Heading 1" Or para.Style = "标题 1" Then
            With para.Range
                .Font.NameFarEast = "宋体"
                .Font.Name = "Times New Roman"
                .Font.Size = 16 ' 小三
                .Font.Bold = True
                .Font.Color = wdColorBlack
                .ParagraphFormat.Alignment = wdAlignParagraphCenter
            End With
        End If
    Next para
    ' MsgBox "一级标题格式化完成！"
End Sub

' 二级标题格式化宏
Sub FormatLevel2Heading()
    Dim para As Paragraph
    For Each para In ActiveDocument.Paragraphs
        If para.Style = "Heading 2" Or para.Style = "标题 2" Then
            With para.Range
                .Font.NameFarEast = "宋体"
                .Font.Name = "Times New Roman"
                .Font.Size = 14 ' 四号
                .Font.Bold = True
                .Font.Color = wdColorBlack
                .ParagraphFormat.Alignment = wdAlignParagraphLeft
            End With
        End If
    Next para
    ' MsgBox "二级标题格式化完成！"
End Sub

' 三级标题格式化宏
Sub FormatLevel3Heading()
    Dim para As Paragraph
    For Each para In ActiveDocument.Paragraphs
        If para.Style = "Heading 3" Or para.Style = "标题 3" Then
            With para.Range
                .Font.NameFarEast = "宋体"
                .Font.Name = "Times New Roman"
                .Font.Size = 12 ' 小四
                .Font.Bold = True
                .Font.Color = wdColorBlack
                .ParagraphFormat.Alignment = wdAlignParagraphLeft
            End With
        End If
    Next para
    ' MsgBox "三级标题格式化完成！"
End Sub

' 正文格式化宏
Sub FormatBodyText()
    Dim para As Paragraph
    For Each para In ActiveDocument.Paragraphs
        ' 假设正文样式为“正文文本”、“Normal”、“First Paragraph”、“正文”
        If para.Style = "正文文本" Or para.Style = "Normal" Or para.Style = "First Paragraph" Or para.Style = "正文" Then
            With para.Range
                .Font.NameFarEast = "宋体"
                .Font.Name = "Times New Roman"
                .Font.Size = 12 ' 小四
                .Font.Bold = False
                .Font.Color = wdColorBlack
                .ParagraphFormat.Alignment = wdAlignParagraphLeft ' 左对齐
                .ParagraphFormat.FirstLineIndent = 24 ' 首行缩进两字符
            End With
        ' Compact样式：字体与正文一样，但无缩进
        ElseIf para.Style = "Compact" Then
            With para.Range
                .Font.NameFarEast = "宋体"
                .Font.Name = "Times New Roman"
                .Font.Size = 12 ' 小四
                .Font.Bold = False
                .Font.Color = wdColorBlack
            End With
        End If
    Next para
    ' MsgBox "正文格式化完成！"
End Sub

Sub SetPageAndBodyFormat()
    ' 设置页面
    With ActiveDocument.PageSetup
        .PaperSize = wdPaperA4
        .TopMargin = CentimetersToPoints(3)
        .BottomMargin = CentimetersToPoints(2.5)
        .LeftMargin = CentimetersToPoints(3)
        .RightMargin = CentimetersToPoints(2.5)
    End With

    ' 设置所有段落行距为1.5倍
    Dim para As Paragraph
    For Each para In ActiveDocument.Paragraphs
        para.Range.ParagraphFormat.LineSpacingRule = wdLineSpace1pt5
    Next para

    ' MsgBox "页面和正文行距设置完成！"
End Sub

' 摘要格式化宏
Sub MergeAndFormatAbstract()
    Dim i As Integer
    Dim para As Paragraph
    Dim nextPara As Paragraph
    Dim txt As String
    Dim contentTxt As String
    Dim rngEnd As Range
    Dim rng As Range
    
    For i = ActiveDocument.Paragraphs.Count To 1 Step -1
        Set para = ActiveDocument.Paragraphs(i)
        txt = Trim(Replace(para.Range.Text, vbCr, ""))
        If txt = "摘要" Or Left(txt, 3) = "摘要：" Or _
           txt = "关键词" Or Left(txt, 4) = "关键词：" Or _
           txt = "Abstract" Or Left(txt, 9) = "Abstract:" Or _
           txt = "Keywords" Or Left(txt, 9) = "Keywords:" Then
            ' 处理摘要、关键词段落
            Dim needMerge As Boolean
            needMerge = (txt = "摘要" Or txt = "关键词" Or txt = "Abstract" Or txt = "Keywords")
            
            If needMerge Then
                ' 需要合并的情况
                ' MsgBox "找到段落内容: [" & para.Range.Text & "]"
                Set nextPara = para.Next
                If Not nextPara Is Nothing Then
                    ' MsgBox "找到内容段: [" & nextPara.Range.Text & "]"
                    contentTxt = nextPara.Range.Text
                    contentTxt = Replace(contentTxt, vbCr, "")
                    contentTxt = Replace(contentTxt, vbLf, "")
                    contentTxt = Trim(contentTxt)
                    ' 如果有下一个标题，只取前面部分
                    Dim nextTitlePos As Integer
                    If txt = "摘要" Or txt = "关键词" Then
                        nextTitlePos = InStr(contentTxt, "Abstract")
                        If nextTitlePos = 0 Then nextTitlePos = InStr(contentTxt, "Keywords")
                    ElseIf txt = "Abstract" Then
                        nextTitlePos = InStr(contentTxt, "Keywords")
                    End If
                    If nextTitlePos > 0 Then
                        contentTxt = Left(contentTxt, nextTitlePos - 1)
                    End If
                    ' 获取 para 段落的最后一个字符（段落符号）前的位置
                    Dim rngInsert As Range
                    Set rngInsert = para.Range.Duplicate
                    rngInsert.End = rngInsert.End - 1  ' 不包括段落符号
                    rngInsert.Collapse wdCollapseEnd
                    rngInsert.InsertAfter contentTxt
                    ' MsgBox "合并后段落内容: [" & para.Range.Text & "]"
                    nextPara.Range.Delete
                Else
                    ' MsgBox "未找到内容段"
                End If
            End If
            
            ' 统一格式化（合并后或直接格式化）
            ' 判断标题后是否有冒号，没有则补冒号
            Dim paraText As String
            paraText = para.Range.Text
            Dim titleLen As Integer
            If Left(txt, 3) = "摘要：" Or Left(txt, 4) = "关键词：" Then
                titleLen = 3
            ElseIf Left(txt, 9) = "Abstract:" Or Left(txt, 9) = "Keywords:" Then
                titleLen = 9
            Else
                titleLen = Len(txt)
            End If
            
            If Len(paraText) < titleLen + 1 Or _
               (Mid(paraText, titleLen + 1, 1) <> "：" And Mid(paraText, titleLen + 1, 1) <> ":") Then
                If txt = "摘要" Or Left(txt, 3) = "摘要：" Then
                    If Len(paraText) < 3 Or Mid(paraText, 3, 1) <> "：" Then
                para.Range.Characters(2).InsertAfter "："
                    End If
                ElseIf txt = "关键词" Or Left(txt, 4) = "关键词：" Then
                    If Len(paraText) < 4 Or Mid(paraText, 4, 1) <> "：" Then
                        para.Range.Characters(3).InsertAfter "："
                    End If
                ElseIf txt = "Abstract" Or Left(txt, 9) = "Abstract:" Then
                    If Len(paraText) < 9 Or Mid(paraText, 9, 1) <> ":" Then
                        para.Range.Characters(8).InsertAfter ":"
                    End If
                ElseIf txt = "Keywords" Or Left(txt, 9) = "Keywords:" Then
                    If Len(paraText) < 9 Or Mid(paraText, 9, 1) <> ":" Then
                        para.Range.Characters(8).InsertAfter ":"
                    End If
                End If
            End If
            
            ' 设置段落样式和字体
            para.Style = ActiveDocument.Styles("正文文本")
            
            ' 根据语言设置字体
            If txt = "摘要" Or Left(txt, 3) = "摘要：" Or txt = "关键词" Or Left(txt, 4) = "关键词：" Then
                ' 中文：宋体
                With para.Range.Font
                    .NameFarEast = "宋体"
                    .Name = "宋体"
                    .Size = 12
                    .Bold = False
                    .Color = wdColorBlack
                End With
            Else
                ' 英文：Times New Roman
                With para.Range.Font
                    .NameFarEast = "Times New Roman"
                    .Name = "Times New Roman"
                    .Size = 12
                    .Bold = False
                    .Color = wdColorBlack
                End With
            End If
            
            ' 设置标题加粗
            If txt = "摘要" Or Left(txt, 3) = "摘要：" Then
                para.Range.Characters(1).Font.Bold = True
                para.Range.Characters(2).Font.Bold = True
                If para.Range.Characters.Count >= 3 Then
                    para.Range.Characters(3).Font.Bold = True
                End If
            ElseIf txt = "关键词" Or Left(txt, 4) = "关键词：" Then
                para.Range.Characters(1).Font.Bold = True
                para.Range.Characters(2).Font.Bold = True
                para.Range.Characters(3).Font.Bold = True
                If para.Range.Characters.Count >= 4 Then
                    para.Range.Characters(4).Font.Bold = True
                End If
            ElseIf txt = "Abstract" Or Left(txt, 9) = "Abstract:" Then
                For j = 1 To 8
                    para.Range.Characters(j).Font.Bold = True
                Next j
                If para.Range.Characters.Count >= 9 Then
                    para.Range.Characters(9).Font.Bold = True
                End If
            ElseIf txt = "Keywords" Or Left(txt, 9) = "Keywords:" Then
                For k = 1 To 8
                    para.Range.Characters(k).Font.Bold = True
                Next k
                If para.Range.Characters.Count >= 9 Then
                    para.Range.Characters(9).Font.Bold = True
                End If
            End If
            
            ' 最后设置段落缩进
            With para.Range.ParagraphFormat
                .Alignment = wdAlignParagraphLeft
                .FirstLineIndent = 24 ' 首行缩进两字符
            End With
        End If
    Next i
    ' MsgBox "摘要格式化完成！"


End Sub

' 目录处理相关宏

' 查找目录位置
Sub FindTableOfContentsPosition()
    Dim para As Paragraph
    For Each para In ActiveDocument.Paragraphs
        If Trim(Replace(para.Range.Text, vbCr, "")) = "目录" Then
            ' MsgBox "找到目录位置：" & para.Range.Start
            Exit Sub
        End If
    Next para
    ' MsgBox "未找到目录标记"
End Sub

' 插入目录（幂等实现）
Sub InsertTableOfContents()
    Dim para As Paragraph
    Dim tocPara As Paragraph
    Dim tocRange As Range
    Dim found As Boolean
    Dim tocTitlePara As Paragraph
    Dim tocTitleRange As Range
    Dim i As Integer
    Dim fld As Field
    found = False
    Set tocPara = Nothing

    ' 1. 找到第一个“目录”标题段落
    For i = 1 To ActiveDocument.Paragraphs.Count
        Set para = ActiveDocument.Paragraphs(i)
        If Trim(Replace(para.Range.Text, vbCr, "")) = "目录" Then
            Set tocPara = para
            Exit For
        End If
    Next i
    If tocPara Is Nothing Then
        MsgBox "未找到目录标记，请在文档中插入'目录'段落"
        Exit Sub
    End If
    ' 新增：如果目录段落样式已经是TOC 标题，则直接跳过
    If tocPara.Style = "TOC 标题" Then
        Exit Sub
    End If

    ' 获取tocPara的索引
    Dim idx As Integer
    idx = 0
    For i = 1 To ActiveDocument.Paragraphs.Count
        If ActiveDocument.Paragraphs(i) Is tocPara Then
            idx = i - 1
            Exit For
        End If
    Next i

    ' 获取tocPara的Range
    Dim cleanRange As Range
    Set cleanRange = tocPara.Range.Duplicate
    cleanRange.Collapse wdCollapseStart

    ' 向上查找并删除所有连续分页符
    Do While cleanRange.Start > 0
        cleanRange.MoveStart wdParagraph, -1
        Dim txt As String
        txt = cleanRange.Paragraphs(1).Range.Text
        If Trim(txt) = Chr(12) Or Trim(txt) = Chr(12) & vbCr Then
            cleanRange.Paragraphs(1).Range.Delete
        ElseIf Right(txt, 1) = Chr(12) Then
            ' 删除段落末尾分页符
            Dim delRange As Range
            Set delRange = cleanRange.Paragraphs(1).Range.Duplicate
            delRange.End = delRange.End - 1
            delRange.Characters(delRange.Characters.Count).Delete
        Else
            Exit Do
        End If
        cleanRange.Collapse wdCollapseStart
    Loop

    ' 2. 在tocPara前插入分页符和新目录标题
    Set tocRange = tocPara.Range.Duplicate
    tocRange.Collapse wdCollapseStart
    tocRange.InsertBreak Type:=wdPageBreak
    tocRange.InsertAfter "目录" & vbCr
    tocRange.Collapse wdCollapseEnd

    ' 3. 删除所有“目录”标题段落（除了刚插入的那个）
    For i = ActiveDocument.Paragraphs.Count To 1 Step -1
        Set para = ActiveDocument.Paragraphs(i)
        If Trim(Replace(para.Range.Text, vbCr, "")) = "目录" And para.Range.Start <> tocRange.Start Then
            para.Range.Delete
        End If
    Next i

    ' 4. 删除所有目录域（wdFieldTOC）
    For Each fld In ActiveDocument.Fields
        If fld.Type = wdFieldTOC Then
            fld.Delete
        End If
    Next fld

    ' 5. 设置TOC 标题样式
    For i = 1 To ActiveDocument.Paragraphs.Count
        Set para = ActiveDocument.Paragraphs(i)
        If Trim(Replace(para.Range.Text, vbCr, "")) = "目录" Then
            Set tocTitlePara = para
            Set tocTitleRange = para.Range.Duplicate
            Exit For
        End If
    Next i
    ApplyTOCTitleStyle tocTitlePara

    ' 6. 插入目录域
    tocTitleRange.Collapse wdCollapseEnd
    tocTitleRange.Fields.Add Range:=tocTitleRange, Type:=wdFieldTOC, Text:="\o ""1-3"" \h \z \u", PreserveFormatting:=True
    tocTitleRange.Fields.Update
    ' 找到目录域的结束位置，在其后插入分页符
    For Each fld In ActiveDocument.Fields
        If fld.Type = wdFieldTOC Then
            Dim endRange As Range
            Set endRange = ActiveDocument.Range(fld.Result.End, fld.Result.End)
            endRange.InsertBreak Type:=wdPageBreak
            Exit For
        End If
    Next fld
    ' 7. 格式化目录条目
    FormatTableOfContentsEntries
End Sub

' 更新目录
Sub UpdateTableOfContents()
    Dim fld As Field
    Dim updated As Boolean
    
    updated = False
    For Each fld In ActiveDocument.Fields
        If fld.Type = wdFieldTOC Then
            fld.Update
            updated = True
        End If
    Next fld
    
    If updated Then
        ' MsgBox "目录更新完成！"
    Else
        ' MsgBox "未找到目录域，请先插入目录"
    End If
End Sub

' 查找文档中的第一个目录域
Private Function GetFirstTOCField() As Field
    Dim fld As Field
    
    For Each fld In ActiveDocument.Fields
        If fld.Type = wdFieldTOC Then
            Set GetFirstTOCField = fld
            Exit Function
        End If
    Next fld
End Function

Private Sub ConfigureTOCTitleStyle()
    Dim tocStyle As Style
    
    On Error Resume Next
    Set tocStyle = ActiveDocument.Styles("TOC 标题")
    On Error GoTo 0
    
    If tocStyle Is Nothing Then Exit Sub
    
    With tocStyle.Font
        .NameFarEast = "宋体"
        .Name = "Times New Roman"
        .Size = 18
        .Bold = True
        .Color = wdColorBlack
    End With
    
    With tocStyle.ParagraphFormat
        .Alignment = wdAlignParagraphCenter
        .FirstLineIndent = 0
        .LeftIndent = 0
        .RightIndent = 0
        .LineSpacingRule = wdLineSpace1pt5
    End With
End Sub

Private Sub ApplyTOCTitleStyle(ByVal para As Paragraph)
    ConfigureTOCTitleStyle
    
    On Error Resume Next
    para.Style = ActiveDocument.Styles("TOC 标题")
    If Err.Number <> 0 Then
        Err.Clear
        para.Style = ActiveDocument.Styles("正文文本")
        If Err.Number <> 0 Then
            Err.Clear
            para.Style = ActiveDocument.Styles("Normal")
        End If

        With para.Range.Font
            .NameFarEast = "宋体"
            .Name = "Times New Roman"
            .Size = 18
            .Bold = True
            .Color = wdColorBlack
        End With
        
        With para.Range.ParagraphFormat
            .Alignment = wdAlignParagraphCenter
            .FirstLineIndent = 0
        End With
    End If
    On Error GoTo 0
End Sub

' 如果已存在目录域，则统一目录标题并更新目录
Private Function NormalizeExistingTableOfContents() As Boolean
    Dim tocField As Field
    Dim tocPara As Paragraph
    Dim tocLastPara As Paragraph
    Dim titlePara As Paragraph
    Dim oldTitlePara As Paragraph
    Dim nextPara As Paragraph
    Dim insertRange As Range
    Dim headingText As String
    Dim paraIndex As Integer
    Dim tocLastParaIndex As Integer
    Dim i As Integer
    Dim nextIndex As Integer
    Dim nextText As String
    
    Set tocField = GetFirstTOCField()
    If tocField Is Nothing Then Exit Function
    
    Set tocPara = tocField.Result.Paragraphs(1)
    Set tocLastPara = tocField.Result.Paragraphs(tocField.Result.Paragraphs.Count)
    Set titlePara = Nothing
    Set oldTitlePara = Nothing
    Set nextPara = Nothing
    paraIndex = 0
    tocLastParaIndex = 0
    
    For i = 1 To ActiveDocument.Paragraphs.Count
        If ActiveDocument.Paragraphs(i).Range.Start = tocPara.Range.Start Then
            paraIndex = i
            Exit For
        End If
    Next i
    
    ' 优先复用目录前的独立标题段，避免直接改写目录域所在段落
    If tocPara.Range.Start > 0 And paraIndex > 1 Then
        Set titlePara = ActiveDocument.Paragraphs(paraIndex - 1)
        headingText = Trim(Replace(titlePara.Range.Text, vbCr, ""))
        
        ' 如果前一段已经是目录条目的一部分，或者本身为空，就在目录前新插入标题段
        If Len(headingText) = 0 Or Left(headingText, 1) = Chr(9) Then
            Set titlePara = Nothing
        ElseIf headingText <> "目录" Then
            Set oldTitlePara = titlePara
            Set titlePara = Nothing
        End If
    End If
    
    If titlePara Is Nothing Then
        Set insertRange = ActiveDocument.Range(tocPara.Range.Start, tocPara.Range.Start)
        insertRange.InsertBefore "目录" & vbCr
        Set titlePara = ActiveDocument.Range(insertRange.Start, insertRange.Start).Paragraphs(1)
        
        If Not oldTitlePara Is Nothing Then
            oldTitlePara.Range.Delete
        End If
    End If
    
    ApplyTOCTitleStyle titlePara
    EnsurePageBreakBeforeParagraph titlePara
    
    For i = 1 To ActiveDocument.Paragraphs.Count
        If ActiveDocument.Paragraphs(i).Range.Start = tocLastPara.Range.Start Then
            tocLastParaIndex = i
            Exit For
        End If
    Next i
    
    tocField.Update
    
    FormatTableOfContentsEntries
    NormalizeExistingTableOfContents = True
End Function

' 设置目录条目格式（不加粗）
Sub FormatTableOfContentsEntries()
    Dim para As Paragraph
    Dim txt As String
    Dim tocFound As Boolean
    
    tocFound = False
    
    For Each para In ActiveDocument.Paragraphs
        txt = Trim(Replace(para.Range.Text, vbCr, ""))
        
        ' 检查是否到达目录部分
        If txt = "目录" Then
            tocFound = True
            GoTo NextPara
        End If
        
        ' 检查是否到达下一个标题（结束目录部分）
        If tocFound And (para.Style = "标题 1" Or para.Style = "标题 2" Or para.Style = "标题 3" Or _
           para.Style = "Heading 1" Or para.Style = "Heading 2" Or para.Style = "Heading 3") Then
            tocFound = False
            GoTo NextPara
        End If
        
        ' 如果在目录部分，设置条目格式
        If tocFound And Len(txt) > 0 And txt <> "目录" Then
            ' 设置目录条目不加粗
            With para.Range.Font
                .Bold = False
                .NameFarEast = "宋体"
                .Name = "Times New Roman"
                .Size = 12
                .Color = wdColorBlack
            End With
        End If
        
NextPara:
    Next para
End Sub

' 设置目录格式（已废弃，避免样式继承问题）
' Sub FormatTableOfContents()
'     ' 此函数已被移除，避免样式继承问题
' End Sub

' 完整的目录处理（查找位置、插入目录、设置格式）
Sub ProcessTableOfContents()
    ' 如果已经存在目录域，则直接统一标题、更新并格式化
    If NormalizeExistingTableOfContents() Then
        Exit Sub
    End If
    
    ' 先查找目录位置
    FindTableOfContentsPosition
    ' 插入目录
    InsertTableOfContents
    ' 更新目录
    UpdateTableOfContents
    ' MsgBox "目录处理完成！"
End Sub



' 参考文献格式化宏
Sub FormatReferences()
    Dim para As Paragraph
    Dim txt As String
    Dim i As Integer
    
    For i = 1 To ActiveDocument.Paragraphs.Count
        Set para = ActiveDocument.Paragraphs(i)
        txt = Trim(Replace(para.Range.Text, vbCr, ""))
        
        ' 查找参考文献标题
        If txt = "参考文献" Or txt = "References" Or _
           Left(txt, 5) = "参考文献：" Or Left(txt, 11) = "References:" Then
            ' 先设置样式，再格式化
            On Error Resume Next
            para.Style = ActiveDocument.Styles("标题 1")
            If Err.Number <> 0 Then
                ' 如果标题1样式不存在，尝试使用默认标题样式
                para.Style = ActiveDocument.Styles("Heading 1")
            End If
            On Error GoTo 0
            
            ' 强制应用格式（覆盖样式）
            With para.Range.Font
                .NameFarEast = "宋体"
                .Name = "宋体"
                .Size = 18 ' 小二
                .Bold = True
                .Color = wdColorBlack
            End With
            With para.Range.ParagraphFormat
                .Alignment = wdAlignParagraphCenter
                .FirstLineIndent = 0
                .LeftIndent = 0
                .RightIndent = 0
            End With
            EnsurePageBreakBeforeParagraph para
            
        End If
    Next i
    
    ' MsgBox "参考文献标题格式化完成！"
End Sub

' 格式化参考文献条目
Sub FormatReferenceEntries()
    Dim para As Paragraph
    Dim txt As String
    Dim i As Integer
    Dim foundReferences As Boolean
    Dim referenceCount As Integer
    Dim endRange As Range
    Dim checkRange As Range
    
    foundReferences = False
    referenceCount = 0
    
    For i = 1 To ActiveDocument.Paragraphs.Count
        Set para = ActiveDocument.Paragraphs(i)
        txt = Trim(Replace(para.Range.Text, vbCr, ""))
        
        ' 检查是否到达参考文献部分
        If txt = "参考文献" Or txt = "References" Or _
           Left(txt, 5) = "参考文献：" Or Left(txt, 11) = "References:" Then
            foundReferences = True
            GoTo NextPara
        End If
        
        ' 检查是否到达其他章节（结束参考文献部分）
        If foundReferences And (txt = "附录" Or txt = "Appendix" Or _
           Left(txt, 3) = "图 " Or Left(txt, 3) = "表 " Or _
           Left(txt, 4) = "Figure" Or Left(txt, 4) = "Table" Or _
           Left(txt, 5) = "致谢" Or Left(txt, 5) = "Acknowledgments" Or _
           Left(txt, 6) = "作者简介" Or Left(txt, 6) = "Author Bio") Then
            foundReferences = False
            GoTo NextPara
        End If
        
        ' 通用判断：检查是否遇到下一个标题样式（结束参考文献部分）
        If foundReferences And (para.Style = "标题 1" Or para.Style = "标题 2" Or para.Style = "标题 3" Or _
           para.Style = "Heading 1" Or para.Style = "Heading 2" Or para.Style = "Heading 3") Then
            ' 在参考文献部分结束后添加分页符
            Set endRange = para.Range.Duplicate
            endRange.Collapse wdCollapseStart
            If endRange.Start > 0 Then
                Set checkRange = ActiveDocument.Range(endRange.Start - 1, endRange.Start)
                If checkRange.Text <> Chr(12) Then
                    endRange.InsertBreak Type:=wdPageBreak
                End If
            End If
            foundReferences = False
            GoTo NextPara
        End If
        
        ' 如果在参考文献部分，格式化条目
        If foundReferences Then
            ' 跳过空行
            If Len(txt) = 0 Then
                GoTo NextPara
            End If
            
            ' 检查是否为参考文献条目（不是空行且不是标题，且不是其他章节标题）
            If Len(txt) > 0 And txt <> "参考文献" And txt <> "References" And _
               Left(txt, 5) <> "参考文献：" And Left(txt, 11) <> "References:" And _
               Left(txt, 3) <> "图 " And Left(txt, 3) <> "表 " And _
               Left(txt, 4) <> "Figure" And Left(txt, 4) <> "Table" And _
               Left(txt, 5) <> "致谢" And Left(txt, 5) <> "Acknowledgments" And _
               Left(txt, 6) <> "作者简介" And Left(txt, 6) <> "Author Bio" And _
               txt <> "附录" And txt <> "Appendix" And _
               para.Style <> "标题 1" And para.Style <> "标题 2" And para.Style <> "标题 3" And _
               para.Style <> "Heading 1" And para.Style <> "Heading 2" And para.Style <> "Heading 3" Then
                referenceCount = referenceCount + 1
                
                ' 先设置段落格式（悬挂缩进）
                With para.Range.ParagraphFormat
                    .Alignment = wdAlignParagraphLeft
                    .FirstLineIndent = -36 ' 首行缩进为负值，实现悬挂缩进（APA标准：0.5英寸）
                    .LeftIndent = 36 ' 左缩进0.5英寸（APA标准）
                    .LineSpacingRule = wdLineSpace1pt5
                End With
                
                ' 再设置字体格式
                With para.Range.Font
                    .NameFarEast = "宋体"
                    .Name = "Times New Roman"
                    .Size = 12 ' 小四
                    .Bold = False
                    .Color = wdColorBlack
                End With
            End If
        End If
        
NextPara:
    Next i
    
    ' MsgBox "参考文献条目格式化完成！共处理 " & referenceCount & " 个条目。"
End Sub

' 自动编号参考文献
Sub AutoNumberReferences()
    Dim para As Paragraph
    Dim txt As String
    Dim i As Integer
    Dim foundReferences As Boolean
    Dim referenceCount As Integer
    Dim newText As String
    
    foundReferences = False
    referenceCount = 0
    
    For i = 1 To ActiveDocument.Paragraphs.Count
        Set para = ActiveDocument.Paragraphs(i)
        txt = Trim(Replace(para.Range.Text, vbCr, ""))
        
        ' 检查是否到达参考文献部分
        If txt = "参考文献" Or txt = "References" Or _
           Left(txt, 5) = "参考文献：" Or Left(txt, 11) = "References:" Then
            foundReferences = True
            GoTo NextPara2
        End If
        
        ' 如果在参考文献部分，处理条目
        If foundReferences Then
            ' 跳过空行
            If Len(txt) = 0 Then
                GoTo NextPara2
            End If
            
            ' 检查是否为参考文献条目（不以数字开头，且不是标题）
            If Not IsNumeric(Left(txt, 1)) And Left(txt, 1) <> "[" And Left(txt, 1) <> "(" And _
               txt <> "参考文献" And txt <> "References" And _
               Left(txt, 5) <> "参考文献：" And Left(txt, 11) <> "References:" Then
                
                referenceCount = referenceCount + 1
                newText = "[" & referenceCount & "] " & txt
                
                ' 替换段落内容
                para.Range.Text = newText & vbCr
            End If
        End If
        
NextPara2:
    Next i
    
    ' MsgBox "参考文献自动编号完成！共编号 " & referenceCount & " 个条目。"
End Sub

' 完整的参考文献处理宏（APA格式）
Sub ProcessReferences()
    ' 1. 格式化参考文献标题
    FormatReferences
    ' 2. 格式化参考文献条目
    FormatReferenceEntries
    
    ' MsgBox "参考文献处理完成！"
End Sub

' 参考文献按字母排序宏
Sub SortReferences()
    Dim para As Paragraph
    Dim txt As String
    Dim i As Integer
    Dim foundReferences As Boolean
    Dim referenceCount As Integer
    Dim references() As String
    Dim referenceRanges() As Range
    Dim tempText As String
    Dim tempRange As Range
    Dim j As Integer, k As Integer
    Dim endPageRange As Range
    Dim checkRange As Range
    
    foundReferences = False
    referenceCount = 0
    ReDim references(0)
    ReDim referenceRanges(0)
    
    ' 第一步：收集参考文献条目
    For i = 1 To ActiveDocument.Paragraphs.Count
        Set para = ActiveDocument.Paragraphs(i)
        txt = Trim(Replace(para.Range.Text, vbCr, ""))
        
        ' 检查是否到达参考文献部分
        If txt = "参考文献" Or txt = "References" Or _
           Left(txt, 5) = "参考文献：" Or Left(txt, 11) = "References:" Then
            foundReferences = True
            GoTo NextParaSort
        End If
        
        ' 通用判断：检查是否遇到下一个标题样式（结束参考文献部分）
        If foundReferences And (para.Style = "标题 1" Or para.Style = "标题 2" Or para.Style = "标题 3" Or _
           para.Style = "Heading 1" Or para.Style = "Heading 2" Or para.Style = "Heading 3") Then
            foundReferences = False
            GoTo NextParaSort
        End If
        
        ' 如果在参考文献部分，收集条目
        If foundReferences Then
            ' 跳过空行
            If Len(txt) = 0 Then
                GoTo NextParaSort
            End If
            
            ' 检查是否为参考文献条目
            If Len(txt) > 0 And txt <> "参考文献" And txt <> "References" And _
               Left(txt, 5) <> "参考文献：" And Left(txt, 11) <> "References:" And _
               Left(txt, 3) <> "图 " And Left(txt, 3) <> "表 " And _
               Left(txt, 4) <> "Figure" And Left(txt, 4) <> "Table" And _
               Left(txt, 5) <> "致谢" And Left(txt, 5) <> "Acknowledgments" And _
               Left(txt, 6) <> "作者简介" And Left(txt, 6) <> "Author Bio" And _
               txt <> "附录" And txt <> "Appendix" And _
               para.Style <> "标题 1" And para.Style <> "标题 2" And para.Style <> "标题 3" And _
               para.Style <> "Heading 1" And para.Style <> "Heading 2" And para.Style <> "Heading 3" Then
                
                referenceCount = referenceCount + 1
                ReDim Preserve references(referenceCount - 1)
                ReDim Preserve referenceRanges(referenceCount - 1)
                
                references(referenceCount - 1) = txt
                Set referenceRanges(referenceCount - 1) = para.Range.Duplicate
            End If
        End If
        
NextParaSort:
    Next i
    
    ' 第二步：按字母排序（不区分大小写，符合APA格式）
    For j = 0 To referenceCount - 2
        For k = j + 1 To referenceCount - 1
            If LCase(references(j)) > LCase(references(k)) Then
                ' 交换文本
                tempText = references(j)
                references(j) = references(k)
                references(k) = tempText
                
                ' 交换范围
                Set tempRange = referenceRanges(j)
                Set referenceRanges(j) = referenceRanges(k)
                Set referenceRanges(k) = tempRange
            End If
        Next k
    Next j
    
    ' 第三步：重新排列段落
    If referenceCount > 0 Then
        ' 删除所有参考文献条目
        For j = 0 To referenceCount - 1
            referenceRanges(j).Delete
        Next j
        
        ' 找到参考文献标题位置
        Dim insertRange As Range
        For i = 1 To ActiveDocument.Paragraphs.Count
            Set para = ActiveDocument.Paragraphs(i)
            txt = Trim(Replace(para.Range.Text, vbCr, ""))
            
            If txt = "参考文献" Or txt = "References" Or _
               Left(txt, 5) = "参考文献：" Or Left(txt, 11) = "References:" Then
                Set insertRange = para.Range.Duplicate
                insertRange.Collapse wdCollapseEnd
                Exit For
            End If
        Next i
        
        ' 按排序后的顺序插入
        For j = 0 To referenceCount - 1
            insertRange.InsertAfter references(j) & vbCr
            ' 确保插入的段落使用正文样式
            Dim newPara As Paragraph
            Set newPara = insertRange.Paragraphs(insertRange.Paragraphs.Count)
            If Not newPara Is Nothing Then
                On Error Resume Next
                newPara.Style = ActiveDocument.Styles("正文文本")
                If Err.Number <> 0 Then
                    ' 如果正文文本样式不存在，尝试使用默认样式
                    newPara.Style = ActiveDocument.Styles("Normal")
                End If
                On Error GoTo 0
            End If
        Next j
        
        ' 在参考文献部分结束后添加分页符
        Dim lastRefPara As Paragraph
        Set lastRefPara = insertRange.Paragraphs(insertRange.Paragraphs.Count)
        If Not lastRefPara Is Nothing Then
            Set endPageRange = lastRefPara.Range.Duplicate
            endPageRange.Collapse wdCollapseEnd
            If endPageRange.Start < ActiveDocument.Content.End Then
                Set checkRange = ActiveDocument.Range(endPageRange.Start, endPageRange.Start + 1)
                If checkRange.Text <> Chr(12) Then
                    endPageRange.InsertBreak Type:=wdPageBreak
                End If
            Else
                endPageRange.InsertBreak Type:=wdPageBreak
            End If
        End If
    End If
    
    ' MsgBox "参考文献排序完成！共排序 " & referenceCount & " 个条目。"
End Sub

' 完整的参考文献处理宏（包含排序）
Sub ProcessReferencesWithSort()
    ' 1. 格式化参考文献标题
    FormatReferences
    ' 2. 排序参考文献条目
    SortReferences
    ' 3. 格式化参考文献条目
    FormatReferenceEntries
    
    ' MsgBox "参考文献处理完成（包含排序）！"
End Sub

' 对外入口：全文使用阿拉伯数字页码
Public Sub ApplyArabicPageNumbersOnly()
    ApplyPageNumbers PAGE_NUMBER_MODE_ARABIC_ONLY
End Sub

' 对外入口：目录前使用罗马数字，目录后使用阿拉伯数字
Public Sub ApplyMixedPageNumbersByTOC()
    ApplyPageNumbers PAGE_NUMBER_MODE_MIXED_BY_TOC
End Sub

Private Sub ApplyPageNumbers(ByVal mode As Integer)
    ClearAllPageNumbers
    
    Select Case mode
        Case PAGE_NUMBER_MODE_ARABIC_ONLY
            ApplyArabicPageNumbersToAllSections
        Case PAGE_NUMBER_MODE_MIXED_BY_TOC
            EnsureSectionBreakAfterTableOfContents
            ApplyMixedPageNumbersBySections
        Case Else
            Err.Raise vbObjectError + 1000, , "未知的页码模式。"
    End Select
End Sub

Private Sub EnsureSectionBreakAfterTableOfContents()
    Dim tocField As Field
    Dim para As Paragraph
    Dim cleanupPara As Paragraph
    Dim nextPara As Paragraph
    Dim breakRange As Range
    Dim gapRange As Range
    Dim txt As String
    Dim i As Integer
    
    If ActiveDocument.Sections.Count > 1 Then Exit Sub
    
    Set tocField = GetFirstTOCField()
    If tocField Is Nothing Then Exit Sub

    Set nextPara = Nothing
    For Each para In ActiveDocument.Paragraphs
        If para.Range.Start >= tocField.Result.End Then
            txt = Trim(Replace(para.Range.Text, vbCr, ""))
            If Len(txt) > 0 Then
                Set nextPara = para
                Exit For
            End If
        End If
    Next para

    If nextPara Is Nothing Then Exit Sub

    Set gapRange = ActiveDocument.Range(tocField.Result.End, nextPara.Range.Start)
    With gapRange.Find
        .ClearFormatting
        .Text = "^b"
        .Forward = True
        .Wrap = wdFindStop
        If .Execute Then
            Exit Sub
        End If
    End With

    If GetSectionIndexByPosition(nextPara.Range.Start) > GetSectionIndexByPosition(tocField.Result.Start) Then
        Exit Sub
    End If

    For i = ActiveDocument.Paragraphs.Count To 1 Step -1
        Set cleanupPara = ActiveDocument.Paragraphs(i)
        If cleanupPara.Range.Start >= tocField.Result.End And cleanupPara.Range.Start < nextPara.Range.Start Then
            txt = Trim(Replace(cleanupPara.Range.Text, vbCr, ""))
            If Len(txt) = 0 Then
                On Error Resume Next
                cleanupPara.Range.Delete
                On Error GoTo 0
            End If
        End If
    Next i

    nextPara.Range.ParagraphFormat.PageBreakBefore = False
    Set breakRange = nextPara.Range.Duplicate
    breakRange.Collapse wdCollapseStart
    breakRange.InsertBreak Type:=wdSectionBreakNextPage
End Sub

Private Function GetSectionIndexByPosition(ByVal pos As Long) As Integer
    Dim i As Integer
    
    For i = 1 To ActiveDocument.Sections.Count
        If pos >= ActiveDocument.Sections(i).Range.Start And pos <= ActiveDocument.Sections(i).Range.End Then
            GetSectionIndexByPosition = i
            Exit Function
        End If
    Next i
End Function

Private Sub ClearAllPageNumbers()
    Dim sec As Section
    Dim hf As HeaderFooter
    Dim idx As Integer
    
    For Each sec In ActiveDocument.Sections
        For idx = wdHeaderFooterPrimary To wdHeaderFooterEvenPages
            Set hf = sec.Footers(idx)
            On Error Resume Next
            hf.LinkToPrevious = False
            hf.PageNumbers.RestartNumberingAtSection = False
            hf.PageNumbers.NumberStyle = wdPageNumberStyleArabic
            Do While hf.PageNumbers.Count > 0
                hf.PageNumbers(1).Delete
            Loop
            On Error GoTo 0
        Next idx
    Next sec
End Sub

Private Sub EnsureCenteredFooterPageNumber(ByVal sec As Section)
    With sec.Footers(wdHeaderFooterPrimary)
        .LinkToPrevious = False
        If .PageNumbers.Count = 0 Then
            .PageNumbers.Add PageNumberAlignment:=wdAlignPageNumberCenter, FirstPage:=True
        End If
    End With
End Sub

Private Sub ApplyArabicPageNumbersToAllSections()
    Dim sec As Section
    
    For Each sec In ActiveDocument.Sections
        EnsureCenteredFooterPageNumber sec
        With sec.Footers(wdHeaderFooterPrimary).PageNumbers
            .NumberStyle = wdPageNumberStyleArabic
            .RestartNumberingAtSection = False
        End With
    Next sec
End Sub

Private Sub ApplyMixedPageNumbersBySections()
    Dim tocField As Field
    Dim tocSectionIndex As Integer
    Dim i As Integer
    
    Set tocField = GetFirstTOCField()
    If tocField Is Nothing Then
        ApplyArabicPageNumbersToAllSections
        Exit Sub
    End If
    
    tocSectionIndex = GetSectionIndexByPosition(tocField.Result.End)
    If tocSectionIndex = 0 Then
        ApplyArabicPageNumbersToAllSections
        Exit Sub
    End If
    
    For i = 1 To ActiveDocument.Sections.Count
        EnsureCenteredFooterPageNumber ActiveDocument.Sections(i)
        
        With ActiveDocument.Sections(i).Footers(wdHeaderFooterPrimary).PageNumbers
            If i <= tocSectionIndex Then
                .NumberStyle = wdPageNumberStyleLowercaseRoman
                .RestartNumberingAtSection = False
            Else
                .NumberStyle = wdPageNumberStyleArabic
                If i = tocSectionIndex + 1 Then
                    .RestartNumberingAtSection = True
                    .StartingNumber = 1
                Else
                    .RestartNumberingAtSection = False
                End If
            End If
        End With
    Next i
End Sub

' 对外公开入口：将文档格式化为山东中医药大学论文格式
Public Sub FormatThesisToSDUTCM()
    Dim response As Integer
    
    ' 询问用户是否继续
    response = MsgBox("即将执行山东中医药大学论文格式化，包括：" & vbCrLf & _
                     "1. 页面设置和正文行距" & vbCrLf & _
                     "2. 标题格式化（题目、一级、二级、三级标题）" & vbCrLf & _
                     "3. 正文格式化" & vbCrLf & _
                     "4. 摘要和关键词格式化" & vbCrLf & _
                     "5. 目录处理（已有目录则更新，否则按“目录”位置插入）" & vbCrLf & _
                     "6. 参考文献格式化（包含排序）" & vbCrLf & _
                     "7. 图片与图题处理" & vbCrLf & _
                     "8. 页码处理（目录前罗马数字，目录后阿拉伯数字）" & vbCrLf & vbCrLf & _
                     "是否继续？", vbYesNo + vbQuestion, "山东中医药大学论文格式化")
    
    If response = vbNo Then
        MsgBox "操作已取消。"
        Exit Sub
    End If
    
    ' 开始执行格式化
    Application.ScreenUpdating = False ' 关闭屏幕更新，提高性能
    
    On Error GoTo ErrorHandler
    
    ' 检查文档是否为空
    If ActiveDocument.Paragraphs.Count = 0 Then
        MsgBox "文档为空，无法执行格式化。", vbExclamation, "警告"
        GoTo CleanUp
    End If
    
    ' 执行统一的内部格式化流程
    RunSDUTCMFormatting
    
    GoTo CleanUp

ErrorHandler:
    MsgBox "格式化过程中出现错误：" & vbCrLf & Err.Description, vbCritical, "错误"

CleanUp:
    ' 恢复设置
    Application.ScreenUpdating = True
End Sub

' 内部统一格式化流程
Private Sub RunSDUTCMFormatting()
    Dim para As Paragraph
    Dim txt As String
    Dim i As Integer
    Dim totalParagraphs As Integer
    
    totalParagraphs = ActiveDocument.Paragraphs.Count
    
    ' 1. Page setup
    SetPageAndBodyFormat
    
    ' 2. Apply base paragraph formatting in one pass
    For i = 1 To totalParagraphs
        Set para = ActiveDocument.Paragraphs(i)
        txt = Trim(Replace(para.Range.Text, vbCr, ""))
        
        If para.Style = ZhTitleStyleName() Then
            FormatTitleParagraph para
        ElseIf para.Style = "Heading 1" Or para.Style = ZhHeadingStyleName(1) Then
            FormatLevel1Paragraph para
        ElseIf para.Style = "Heading 2" Or para.Style = ZhHeadingStyleName(2) Then
            FormatLevel2Paragraph para
        ElseIf para.Style = "Heading 3" Or para.Style = ZhHeadingStyleName(3) Then
            FormatLevel3Paragraph para
        ElseIf para.Style = ZhBodyTextStyleName() Or para.Style = "Normal" Or para.Style = "First Paragraph" Or para.Style = ZhBodyStyleName() Then
            FormatBodyParagraph para
        ElseIf para.Style = "Compact" Then
            FormatCompactParagraph para
        End If
    Next i
    
    ' 3. Abstract and keywords
    MergeAndFormatAbstract
    
    ' 4. Table of contents
    ProcessTableOfContents
    
    ' 5. References
    ProcessReferencesWithSort
    
    ' 6. Images
    ProcessImages
    
    ' 7. Mixed page numbers by TOC
    ApplyMixedPageNumbersByTOC
End Sub

' Return the localized Word title style name: 标题
Private Function ZhTitleStyleName() As String
    ZhTitleStyleName = ChrW(&H6807) & ChrW(&H9898)
End Function

' Return the localized Word heading style name: 标题 1/2/3
Private Function ZhHeadingStyleName(level As Integer) As String
    ZhHeadingStyleName = ZhTitleStyleName() & " " & CStr(level)
End Function

' Return the body-text style name: 正文文本
Private Function ZhBodyTextStyleName() As String
    ZhBodyTextStyleName = ChrW(&H6B63) & ChrW(&H6587) & ChrW(&H6587) & ChrW(&H672C)
End Function

' Return the body style name: 正文
Private Function ZhBodyStyleName() As String
    ZhBodyStyleName = ChrW(&H6B63) & ChrW(&H6587)
End Function

Private Sub EnsurePageBreakBeforeParagraph(ByVal para As Paragraph)
    If para Is Nothing Then Exit Sub
    para.Range.ParagraphFormat.PageBreakBefore = True
End Sub

' Format a title paragraph
Sub FormatTitleParagraph(para As Paragraph)
    With para.Range
        .Font.NameFarEast = "黑体"
        .Font.Name = "黑体"
        .Font.Size = 18 ' 小二
        .Font.Bold = True
        .Font.Color = wdColorBlack
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
    End With
End Sub

' 格式化单个一级标题段落
Sub FormatLevel1Paragraph(para As Paragraph)
    With para.Range
        .Font.NameFarEast = "宋体"
        .Font.Name = "Times New Roman"
        .Font.Size = 16 ' 小三
        .Font.Bold = True
        .Font.Color = wdColorBlack
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
    End With
End Sub

' 格式化单个二级标题段落
Sub FormatLevel2Paragraph(para As Paragraph)
    With para.Range
        .Font.NameFarEast = "宋体"
        .Font.Name = "Times New Roman"
        .Font.Size = 14 ' 四号
        .Font.Bold = True
        .Font.Color = wdColorBlack
        .ParagraphFormat.Alignment = wdAlignParagraphLeft
    End With
End Sub

' 格式化单个三级标题段落
Sub FormatLevel3Paragraph(para As Paragraph)
    With para.Range
        .Font.NameFarEast = "宋体"
        .Font.Name = "Times New Roman"
        .Font.Size = 12 ' 小四
        .Font.Bold = True
        .Font.Color = wdColorBlack
        .ParagraphFormat.Alignment = wdAlignParagraphLeft
    End With
End Sub

' 格式化单个正文段落
Sub FormatBodyParagraph(para As Paragraph)
    With para.Range
        .Font.NameFarEast = "宋体"
        .Font.Name = "Times New Roman"
        .Font.Size = 12 ' 小四
        .Font.Bold = False
        .Font.Color = wdColorBlack
        .ParagraphFormat.Alignment = wdAlignParagraphLeft ' 左对齐
        .ParagraphFormat.FirstLineIndent = 24 ' 首行缩进两字符
        .ParagraphFormat.LineSpacingRule = wdLineSpace1pt5
    End With
End Sub

' 格式化单个Compact段落
Sub FormatCompactParagraph(para As Paragraph)
    With para.Range
        .Font.NameFarEast = "宋体"
        .Font.Name = "Times New Roman"
        .Font.Size = 12 ' 小四
        .Font.Bold = False
        .Font.Color = wdColorBlack
    End With
End Sub

' 图片居中格式化宏
Sub FormatImages()
    Dim shp As Shape
    Dim inlineShape As InlineShape
    Dim imageCount As Integer
    
    imageCount = 0
    
    ' 处理浮动图片（Shape对象）
    For Each shp In ActiveDocument.Shapes
        If shp.Type = msoPicture Or shp.Type = msoLinkedPicture Then
            ' 设置图片居中
            With shp
                .Left = (ActiveDocument.PageSetup.PageWidth - .Width) / 2
                .WrapFormat.Type = wdWrapSquare
                .WrapFormat.Side = wdWrapBoth
            End With
            imageCount = imageCount + 1
        End If
    Next shp
    
    ' 处理内嵌图片（InlineShape对象）
    For Each inlineShape In ActiveDocument.InlineShapes
        ' 设置内嵌图片居中
        With inlineShape.Range.ParagraphFormat
            .Alignment = wdAlignParagraphCenter
        End With
        imageCount = imageCount + 1
    Next inlineShape
    
    ' MsgBox "图片居中格式化完成！共处理 " & imageCount & " 个图片。"
End Sub

' 图片标题格式化宏
Sub FormatImageCaptions()
    Dim para As Paragraph
    Dim txt As String
    Dim captionCount As Integer
    
    captionCount = 0
    
    For Each para In ActiveDocument.Paragraphs
        txt = Trim(Replace(para.Range.Text, vbCr, ""))
        
        ' 查找图片标题（以"图"或"Figure"开头）
        If Left(txt, 2) = "图 " Or Left(txt, 7) = "Figure " Or _
           Left(txt, 3) = "图 " Or Left(txt, 8) = "Figure " Then
            ' 格式化图片标题
            With para.Range.Font
                .NameFarEast = "宋体"
                .Name = "宋体"
                .Size = 12 ' 小四
                .Bold = False
                .Color = wdColorBlack
            End With
            With para.Range.ParagraphFormat
                .Alignment = wdAlignParagraphCenter ' 居中
                .FirstLineIndent = 0 ' 无缩进
                .LineSpacingRule = wdLineSpace1pt5
                End With
            captionCount = captionCount + 1
        End If
    Next para
    
    ' MsgBox "图片标题格式化完成！共处理 " & captionCount & " 个标题。"
End Sub

' 完整的图片处理宏
Sub ProcessImages()
    ' 1. 图片居中
    FormatImages
    ' 2. 图片标题格式化
    FormatImageCaptions
    
    ' MsgBox "图片处理完成！"
End Sub

