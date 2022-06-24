Sub CurrencyNumber()
Dim i As Range, Acell As Cell, CR As Range
On Error Resume Next
Application.ScreenUpdating = False
If Selection.Type = 2 Then
For Each i In Selection.Words
If i Like "####*" = True Then
If i.Next Like "." = True And i.Next(wdWord, 2) Like "#*" = True Then
i.SetRange Start:=i.Start, End:=i.Next(wdWord, 2).End
i = Format(i, "Standard")
Else
i = Format(i, "Standard")
End If
End If
Next i
ElseIf Selection.Type = 5 Then
For Each Acell In Selection.Cells
Set CR = ActiveDocument.Range(Acell.Range.Start, Acell.Range.End - 1)
MsgBox CR
If CR Like "####*" = True Then
If CR Like "####.#*" = True Then
Yn = Format(CR, "Standard")
CR.Text = Nn
Else
Yn = Format(CR, "Standard")
CR.Text = Nn
End If
End If
Next Acell
Else
'MsgBox "您只能选定文本或者表格之一!", vbOK + vbInformation
End If
Application.ScreenUpdating = True
End Sub
Sub Form_Level()
UserForm1.Show


'UserForm1.Show

End Sub
Sub Frm_Font_WIH()
Font_WIH.Show
End Sub

Sub Font_body()
    Selection.ClearFormatting  'clean the format pefore
    Selection.Font.Name = "STFangsong"
    Selection.Font.Size = 16
    Selection.Font.Bold = False 'not bold
    
    Selection.ParagraphFormat.OutlineLevel = wdOutlineLevelBodyText
    Selection.ParagraphFormat.CharacterUnitFirstLineIndent = 2 'the first line
    Selection.ParagraphFormat.LeftIndent = CentimetersToPoints(0) '左缩进为0

    Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify '两端对齐
 
    Selection.Paragraphs.LineSpacingRule = wdLineSpaceExactly   '设置行间距固定值28
    Selection.Paragraphs.LineSpacing = 28 '设置行间距固定值28
End Sub
Sub Font_title()
  

    Selection.Font.Name = "Microsoft YaHei"
    Selection.Font.Size = 22
    Selection.Font.Bold = False '不加粗
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
  
       Selection.Paragraphs.LineSpacingRule = wdLineSpaceExactly   '设置行间距固定值28
 Selection.Paragraphs.LineSpacing = 28 '设置行间距固定值28
End Sub
Sub Font_1()
   
    Selection.Font.Name = "SimHei"
    Selection.Font.Size = 16
    Selection.Font.Bold = False '不加粗
    Selection.ParagraphFormat.OutlineLevel = wdOutlineLevel1
    Selection.ParagraphFormat.CharacterUnitFirstLineIndent = 2 '首行缩进2字符
    Selection.ParagraphFormat.Alignment = wdAlignParagraphThaiJustify
   Selection.Paragraphs.LineSpacingRule = wdLineSpaceExactly   '设置行间距固定值28
 Selection.Paragraphs.LineSpacing = 28 '设置行间距固定值28
End Sub
Sub Font_2()

    Selection.Font.Name = "KaiTi"
    Selection.Font.Size = 16
        Selection.Font.Bold = False '不加粗
    Selection.ParagraphFormat.OutlineLevel = wdOutlineLevel2
        Selection.ParagraphFormat.CharacterUnitFirstLineIndent = 2 '首行缩进2字符
   Selection.Paragraphs.LineSpacingRule = wdLineSpaceExactly   '设置行间距固定值28
 Selection.Paragraphs.LineSpacing = 28 '设置行间距固定值28
End Sub

Sub Font_3()
 
    Selection.Font.Name = "STFangsong"
    Selection.Font.Size = 16
        Selection.Font.Bold = False '不加粗
    Selection.ParagraphFormat.OutlineLevel = wdOutlineLevel3
    Selection.ParagraphFormat.CharacterUnitFirstLineIndent = 2 '首行缩进2字符
   Selection.Paragraphs.LineSpacingRule = wdLineSpaceExactly   '设置行间距固定值28
 Selection.Paragraphs.LineSpacing = 28 '设置行间距固定值28

End Sub
Sub Font_4()

    Selection.Font.Name = "STFangsong"
    Selection.Font.Size = 16
        Selection.Font.Bold = False '不加粗
    Selection.ParagraphFormat.OutlineLevel = wdOutlineLevel4
        Selection.ParagraphFormat.CharacterUnitFirstLineIndent = 2 '首行缩进2字符
   Selection.Paragraphs.LineSpacingRule = wdLineSpaceExactly   '设置行间距固定值28
 Selection.Paragraphs.LineSpacing = 28 '设置行间距固定值28
End Sub
Sub SimSun()
    Selection.Font.Name = "SimSun"
    Selection.Font.Size = 12
    Selection.Font.Bold = False '不加粗
'    Selection.ParagraphFormat.OutlineLevel = wdOutlineLevelBodyText
    Selection.Paragraphs.LineSpacingRule = wdLineSpaceExactly       '设置行间距固定值28
    Selection.Paragraphs.LineSpacing = 28 '设置行间距固定值28



    Selection.ParagraphFormat.LeftIndent = CentimetersToPoints(0) '左缩进为0
    Selection.ParagraphFormat.SpaceAfter = InchesToPoints(0) '段前后间距
    Selection.ParagraphFormat.SpaceBefore = InchesToPoints(0)
    Selection.ParagraphFormat.Alignment = wdAlignParagraphThaiJustify  '两端对齐
End Sub
Sub Level()
UserForm1.Show
End Sub
Sub Format_table()
Dim mytable As Table
For Each mytable In ActiveDocument.Tables
        'Selection.Style = ActiveDocument.Styles("普通表格") '清除表格
        'WordBasic.ClearTableStyle
        mytable.Rows.WrapAroundText = False '取消文字环绕
        
       ' mytable.Range.Editors.Add wdEditorEveryone '选中整个表格
      '   mytable.AutoFitBehavior (wdAutoFitWindow) '根据窗口调整内容
      '  mytable.Rows.HeightRule = wdRowHeightAuto '
         mytable.Rows.Height = CentimetersToPoints(0) '上下居中
        mytable.Range.Cells(1).VerticalAlignment = wdCellAlignVerticalCenter '垂直居中
             
        With mytable
            .TopPadding = CentimetersToPoints(0.08)  '上下间距=0.08，0.08
            .BottomPadding = CentimetersToPoints(0.08) '
            .LeftPadding = CentimetersToPoints(0.19) '左右间距0.19
            .RightPadding = CentimetersToPoints(0.19) '
             .Spacing = 0 ''取消固定行高
             .AllowPageBreaks = True        '允许断行
            .AllowAutoFit = True    '自动适应文字
            .Rows(1).Select
            .Rows.HeadingFormat = True  '行标题重复
             .Rows.Alignment = 1 '设置整个表格在页面中水平居中'水平居中
            
            With mytable.Range
             
              .Font.Size = 10
              .Font.Name = "SimSun"
               .Font.Bold = False '加粗
                .ParagraphFormat.Alignment = 1 '水平居中
                .Cells.VerticalAlignment = 1 '垂直居中
            End With
            '=========================================设置边框=========================================
              With mytable.Borders
              .InsideLineStyle = wdLineStyleSingle '单细实线
              .OutsideLineStyle = wdLineStyleSingle '单细实线
              End With
              '========================================================================================
         End With
           
    
   ' ActiveDocument.SelectAllEditableRanges (wdEditorEveryone) '选中全部表格区域
   ' ActiveDocument.DeleteAllEditableRanges (wdEditorEveryone) '删除
  '  Application.ScreenUpdating = True
 

   Next
End Sub
Sub try()
Selection.Paragraphs.LineSpacingRule = wdLineSpaceExactly
 Selection.Paragraphs.LineSpacing = 28 '设置行间距固定值28
End Sub

   
