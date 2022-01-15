'Copyright Moskalev Dmitry
Option Explicit
Private Sub UserForm_Initialize()
Application.DisplayAlerts = False
Dim XMLHTTP As Object
Dim URL$
URL = "***"
Set XMLHTTP = CreateObject("MSXML2.XMLHTTP")
XMLHTTP.Open "GET", URL, False
On Error Resume Next
XMLHTTP.Send
On Error GoTo 0
If XMLHTTP.Status = 200 Then
ComboBox1.Text = "Choose subject"
ComboBox1.AddItem "Programming fundamentals"
ComboBox1.AddItem "Data structures and algorithms"
ComboBox1.AddItem "Object Oriented Programming"
ComboBox1.AddItem "Informatics"
ComboBox1.AddItem "Other subject"
ComboBox3.AddItem "Male"
ComboBox3.AddItem "Female"
ComboBox3.AddItem "Male and Female"
CommandButton1.MousePointer = fmMousePointerNoDrop
CommandButton1.Locked = True
TextBox10.BorderColor = RGB(225, 140, 120)
ComboBox1.ShowDropButtonWhen = fmShowDropButtonWhenFocus
ComboBox2.ShowDropButtonWhen = fmShowDropButtonWhenFocus
ComboBox3.ShowDropButtonWhen = fmShowDropButtonWhenFocus
Options.PrintProperties = False
Options.PrintXMLTag = False
ComboBox1.SetFocus
Label26.Top = UserForm1.Top
Label26.Left = UserForm1.Left
Label26.Width = UserForm1.Width
Label26.Height = UserForm1.Height
MultiPage1.Pages(0).ScrollHeight = Label27.Width
MultiPage1.Pages(0).ScrollHeight = Label27.Height
MultiPage1.Pages(0).ScrollHeight = Label27.Height
MultiPage1.Pages(1).ScrollHeight = Label29.Width
MultiPage1.Pages(1).ScrollHeight = Label29.Height
Else
Set XMLHTTP = Nothing
MsgBox "No connection!"
Application.Quit SaveChanges:=wdDoNotSaveChanges
End If
Set XMLHTTP = Nothing
End Sub
Private Sub TextBox12_Change()
If (Len(TextBox12.Text) > 0) Then
ComboBox2.ShowDropButtonWhen = fmShowDropButtonWhenAlways
Label8.Visible = False
TextBox10.Value = 10
ComboBox2.Enabled = True
ComboBox2.Clear
ComboBox2.AddItem "Laboratory works"
ComboBox2.AddItem "Practical tasks"
ComboBox2.AddItem "Settlement and graphic tasks"
ComboBox2.AddItem "Course works"
Else
ComboBox2.Enabled = False
TextBox10.Value = 0
TextBox12.SetFocus
Label8.Visible = True
End If
End Sub
Private Sub ComboBox1_Change()
ComboBox1.Style = 2
If ComboBox1.Text = "Programming fundamentals" Then
ComboBox2.Clear
ComboBox2.AddItem "Laboratory works"
ComboBox2.AddItem "Settlement and graphic tasks"
End If

If ComboBox1.Text = "Data structures and algorithms" Then
ComboBox2.Clear
ComboBox2.AddItem "Laboratory works"
ComboBox2.AddItem "Course works"
End If

If ComboBox1.Text = "Object Oriented Programming" Then
ComboBox2.Clear
ComboBox2.AddItem "Laboratory works"
ComboBox2.AddItem "Settlement and graphic tasks"
End If

If ComboBox1.Text = "Informatics" Then
ComboBox2.Clear
ComboBox2.AddItem "Practical tasks"
ComboBox2.AddItem "Settlement and graphic tasks"
End If

If (Len(ComboBox1.Text) = 0 Or ComboBox1.Text = "Choose subject" Or (ComboBox1.Text <> "Programming fundamentals" And ComboBox1.Text <> "Data structures and algorithms" And ComboBox1.Text <> "Object Oriented Programming" And ComboBox1.Text <> "Informatics" And ComboBox1.Text <> "Other subject")) Then
ComboBox1.SetFocus
Label8.Visible = True
Else
Label8.Visible = False
If (ComboBox1.Text <> "Other subject") Then
ComboBox2.Enabled = True
ComboBox2.SetFocus
TextBox10.Value = 10
Else
TextBox12.Enabled = True
TextBox12.Locked = False
TextBox12.Visible = True
ComboBox1.Visible = False
TextBox12.SetFocus
ComboBox1.Visible = False
End If
ComboBox1.Enabled = False
End If

End Sub
Private Sub ComboBox2_Change()
TextBox12.Enabled = False
ComboBox2.Style = 2
If ComboBox2.Text = "Laboratory works" Or ComboBox2.Text = "Practical tasks" Then
TextBox10.Value = 25
Label19.Visible = False
Label20.Visible = False
TextBox9.Visible = False
TextBox10.Top = Val(TextBox10.Top - TextBox9.Height)
CommandButton1.Top = Val(CommandButton1.Top - TextBox9.Height)
Else
Label19.Visible = True
TextBox9.Visible = True
End If

If ComboBox2.Text = "Settlement and graphic tasks" Or ComboBox2.Text = "Course works" Then
TextBox10.Value = 30
Label5.Visible = False
Label15.Visible = False
TextBox5.Visible = False
Label7.Visible = False
Label16.Visible = False
TextBox7.Visible = False
TextBox8.Top = TextBox5.Top
Label18.Top = Label5.Top
Label17.Top = Label15.Top
TextBox9.Top = TextBox7.Top
Label19.Top = Label7.Top
Label20.Top = Label16.Top
TextBox10.Top = Val(TextBox10.Top - TextBox5.Height - TextBox7.Height)
CommandButton1.Top = Val(CommandButton1.Top - TextBox5.Height - TextBox7.Height)
Else
Label5.Visible = True
TextBox5.Visible = True
Label7.Visible = True
TextBox7.Visible = True
End If

If (Len(ComboBox2.Text) = 0 Or ComboBox2.Text = "Type of work" Or (ComboBox2.Text <> "Laboratory works" And ComboBox2.Text <> "Practical tasks" And ComboBox2.Text <> "Settlement and graphic tasks" And ComboBox2.Text <> "Course works")) Then
ComboBox2.SetFocus
Label9.Visible = True
Else
Label9.Visible = False
ComboBox2.Enabled = False
ComboBox3.Enabled = True
ComboBox3.SetFocus
End If

End Sub
Private Sub ComboBox3_Change()
ComboBox3.Style = 2
If ComboBox3.Text = "Male" Then
Label3.Visible = False
Label22.Visible = False
Label31.Visible = False
Label21.Visible = True
End If
If ComboBox3.Text = "Female" Then
Label3.Visible = False
Label21.Visible = False
Label31.Visible = False
Label22.Visible = True
End If
If ComboBox3.Text = "Male and Female" Then
Label3.Visible = False
Label22.Visible = False
Label21.Visible = False
Label31.Visible = True
End If


If Len(ComboBox3.Text) = 0 Or ComboBox3.Text = "Choose gender" Or (ComboBox3.Text <> "Male" And ComboBox3.Text <> "Female" And ComboBox3.Text <> "Male and Female") Then
Label10.Visible = True
ComboBox3.SetFocus
Else
If (ComboBox2.Text = "Laboratory works" Or ComboBox2.Text = "Practical tasks") Then
TextBox10.Value = 40
Else
TextBox10.Value = 50
End If
Label10.Visible = False
ComboBox3.Enabled = False
TextBox2.Enabled = True
TextBox2.SetFocus
End If
End Sub
Private Sub TextBox2_Change()
If (Len(TextBox2.Text) > 0) Then
Label12.Visible = False
 If (ComboBox2.Text = "Laboratory works" Or ComboBox2.Text = "Practical tasks") Then
 TextBox10.Value = 50
 TextBox3.Enabled = True
 Else
 TextBox10.Value = 60
 TextBox3.Enabled = True
 End If
Else
If (ComboBox2.Text = "Laboratory works" Or ComboBox2.Text = "Practical tasks") Then
TextBox10.Value = 40
Else
TextBox10.Value = 50
End If
Label12.Visible = True
TextBox3.Enabled = False
End If
End Sub
Private Sub TextBox3_Change()
TextBox2.Enabled = False
If (Len(TextBox3.Text) > 0) Then
Label13.Visible = False
If (UBound(Split(TextBox3.Text, vbNewLine)) + 1 > 1) Then
Label21.Visible = False
Label22.Visible = False
Label31.Visible = True
Else:
If (ComboBox3.Text = "Male") Then
Label31.Visible = False
Label21.Visible = True
End If
If (ComboBox3.Text = "Female") Then
Label31.Visible = False
Label22.Visible = True
End If
End If
If (ComboBox2.Text = "Laboratory works" Or ComboBox2.Text = "Practical tasks") Then
TextBox10.Value = 60
TextBox4.Enabled = True
Else
TextBox10.Value = 70
TextBox4.Enabled = True
End If
Else
If (ComboBox2.Text = "Laboratory works" Or ComboBox2.Text = "Practical tasks") Then
TextBox10.Value = 50
Else
TextBox10.Value = 60
End If
Label13.Visible = True
TextBox4.Enabled = False
End If
End Sub
Private Sub TextBox4_Change()
TextBox3.Enabled = False
If (Len(TextBox4.Text) > 0) Then
Label14.Visible = False
If (UBound(Split(TextBox4.Text, vbNewLine)) + 1 > 1) Then
Label4.Visible = False
Label30.Visible = True
Else:
Label30.Visible = False
Label4.Visible = True
End If
If (ComboBox2.Text = "Laboratory works" Or ComboBox2.Text = "Practical tasks") Then
TextBox10.Value = 70
TextBox5.Enabled = True
Else
TextBox10.Value = 80
TextBox8.Enabled = True
End If
Else
If (ComboBox2.Text = "Laboratory works" Or ComboBox2.Text = "Practical tasks") Then
TextBox10.Value = 60
Else
TextBox10.Value = 70
End If
Label14.Visible = True
TextBox5.Enabled = False
TextBox8.Enabled = False
End If
End Sub
Private Sub TextBox5_Change()
TextBox4.Enabled = False
If (Len(TextBox5.Text) > 0) Then
Label15.Visible = False
TextBox7.Enabled = True
TextBox10.Value = 80
Else
Label15.Visible = True
TextBox7.Enabled = False
TextBox10.Value = 70
End If
End Sub
Private Sub TextBox7_Change()
TextBox5.Enabled = False
If (Len(TextBox7.Text) > 0) Then
Label16.Visible = False
TextBox8.Enabled = True
TextBox10.Value = 90
Else
Label16.Visible = True
TextBox8.Enabled = False
TextBox10.Value = 80
End If
End Sub
Private Sub TextBox8_Change()
TextBox4.Enabled = False
TextBox7.Enabled = False
If (Len(TextBox8.Text) > 0) Then
Label17.Visible = False
If (ComboBox2.Text = "Laboratory works" Or ComboBox2.Text = "Practical tasks") Then
TextBox10.Value = 100
Else
TextBox10.Value = 90
TextBox9.Enabled = True
End If
Else
If (ComboBox2.Text = "Laboratory works" Or ComboBox2.Text = "Practical tasks") Then
TextBox10.Value = 90
Else
TextBox10.Value = 80
End If
Label17.Visible = True
TextBox9.Enabled = False
End If
End Sub
Private Sub TextBox9_Change()
TextBox8.Enabled = False
If Len(TextBox9.Value) > 0 And Val(TextBox9.Value) >= 0 And Val(TextBox9.Value) <= 100 Then
Label20.Visible = False
TextBox10.Value = 100
Else
Label20.Visible = True
TextBox10.Value = 90
End If
End Sub
Private Sub TextBox10_Change()
If Val(TextBox10.Value) = 100 Then
CommandButton1.Locked = False
TextBox10.BorderColor = vbGreen
CommandButton1.MousePointer = fmMousePointerDefault
Else
CommandButton1.Locked = True
TextBox10.BorderColor = RGB(225, 140, 120)
CommandButton1.MousePointer = fmMousePointerNoDrop
End If
End Sub
Private Sub CommandButton1_Change()
End Sub
Private Sub Label25_Click()
Label26.Visible = True
MultiPage1.Visible = True
CommandButton10.Visible = True
CommandButton10.SetFocus
End Sub
Private Sub CommandButton10_Click()
If (Label26.Visible = True) Then
Label26.Visible = False
MultiPage1.Visible = False
CommandButton10.Visible = False
End If

If (ComboBox1.Text = "Choose subject" And ComboBox1.Enabled = True) Then
ComboBox1.SetFocus
End If

If (ComboBox2.Text = "Type of work" And ComboBox2.Enabled = True) Then
ComboBox2.SetFocus
End If

If (ComboBox3.Text = "Choose gender" And ComboBox3.Enabled = True) Then
ComboBox3.SetFocus
End If

End Sub
Private Sub CommandButton1_Click()
If (Val(TextBox10.Value) = 100 And Label17.Visible = False And Label20.Visible = False And Label26.Visible = False And MultiPage1.Visible = False And CommandButton10.Visible = False And ActiveDocument.ProtectionType = wdAllowOnlyReading) Then
TextBox8.Enabled = False
TextBox9.Enabled = False
CommandButton1.Enabled = False
OptionButton1.Enabled = False
OptionButton2.Enabled = False
CommandButton2.Enabled = False
Application.ScreenUpdating = False
ActiveDocument.Unprotect ("***")
ActiveDocument.Content.Select
Selection.Delete

Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label12.Visible = False
Label13.Visible = False
Label14.Visible = False
Label15.Visible = False
Label16.Visible = False
Label17.Visible = False
Label20.Visible = False

Dim Text1, Text1_2, Text2, Text2_1, Text4, Text5, Text6, Text7 As String

Text1 = "Ministry of Science and Higher Education"
Text1_2 = "Russian Federation"
Text2 = "Federal state budget educational"
Text2_1 = "institution of higher education"
Text4 = "Novosibirsk State Technical University"
Text5 = "Department of Applied Mathematics"
Text6 = "by discipline"
Text7 = "Novosibirsk"

Dim aStory As Range
Dim aField As Field
Dim myTOC As TableOfContents
For Each aStory In ActiveDocument.StoryRanges
For Each aField In aStory.Fields
aField.Update
Next aField
Next aStory
For Each myTOC In ActiveDocument.TablesOfContents
myTOC.Update
Next myTOC
  WordBasic.PageSetupMargins Tab:=3, PaperSize:=9, TopMargin:="2", _
        BottomMargin:="2", LeftMargin:="2", RightMargin:="2", Gutter:="0", _
        PageWidth:="21", PageHeight:="29.7", Orientation:=0, FirstPage:=0, _
        OtherPages:=0, VertAlign:=0, ApplyPropsTo:=4, FacingPages:=0, _
        HeaderDistance:="0.5", FooterDistance:="0.5", SectionStart:=2, _
        OddAndEvenPages:=0, DifferentFirstPage:=1, Endnotes:=0, LineNum:=0, _
        CountBy:=0, TwoOnOne:=0, GutterPosition:=0, LayoutMode:=0, DocFontName:= _
        "", FirstPageOnLeft:=0, SectionType:=1, FolioPrint:=0, ReverseFolio:=0, _
        FolioPages:=1
    Selection.Style = ActiveDocument.Styles("Заголовок 1")
    Selection.Paragraphs(1).SpaceBefore = 0
    Selection.TypeText Text:=Text1 & Chr(10) & Text1_2
    With ActiveDocument.Styles("Обычный").Font
        .Name = "+Основной текст"
        .Size = 12
        .Underline = wdUnderlineNone
        .UnderlineColor = wdColorAutomatic
        .Color = wdColorAutomatic
        .Scaling = 100
        .Kerning = 0
        .Animation = wdAnimationNone
    End With
    Selection.TypeParagraph
    Selection.TypeParagraph
    Selection.Style = ActiveDocument.Styles("Заголовок 2")
    Selection.TypeText Text:=Text2 & Chr(10) & Text2_1
    Selection.TypeParagraph
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.TypeText Text:=Chr(171) & Text4 & Chr(187)
    Selection.MoveLeft Unit:=wdCharacter, Count:=Len(Chr(171) & Text4 & Chr(187)), Extend:=wdExtend
    With Selection.Font
        .Name = "Calibri Light"
        .Size = 16
        .Underline = wdUnderlineNone
        .UnderlineColor = wdColorAutomatic
        .SmallCaps = True
        .Color = 2434341
        .Spacing = 0
        .Scaling = 100
        .Position = 0
        .Kerning = 0
        .Animation = wdAnimationNone
    End With
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeParagraph
    Selection.Font.Size = 12
    Selection.TypeParagraph
    On Error Resume Next
    Selection.InlineShapes.AddPicture FileName:="***", LinkToFile:=False, SaveWithDocument:=True
    On Error GoTo 0
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeParagraph
    Selection.TypeParagraph
    Selection.Style = ActiveDocument.Styles("Заголовок 2")
    Selection.TypeText Text:=Text5
    Selection.TypeParagraph
    Selection.TypeParagraph
    Selection.Style = ActiveDocument.Styles("Заголовок 2")

    Dim Tasknumber As String
    Dim Nameoftype As String
    Tasknumber = TextBox5.Text
    
    If Len(Tasknumber) > 0 Or (TextBox5.Visible = False) Then
        If ComboBox2.Text = "Laboratory works" Then
        Nameoftype = "Laboratory work №"
        Selection.TypeText Text:=Nameoftype
        Selection.TypeText Text:=Tasknumber
        End If
        If ComboBox2.Text = "Practical tasks" Then
        Nameoftype = "Practical task №"
        Selection.TypeText Text:=Nameoftype
        Selection.TypeText Text:=Tasknumber
        End If
        If ComboBox2.Text = "Settlement and graphic tasks" Then
        Nameoftype = "Settlement and graphic task"
        Selection.TypeText Text:=Nameoftype
        End If
        If ComboBox2.Text = "Course works" Then
        Nameoftype = "Course work"
        Selection.TypeText Text:=Nameoftype
        End If
    End If
    Selection.TypeParagraph
    Selection.Style = ActiveDocument.Styles("Заголовок 2")
    Selection.TypeText Text:=Text6 & " " & Chr(171)
    Dim Subject As String
    Subject = ComboBox1.Text & TextBox12.Text
    If Len(Subject) > 0 Then
        If ComboBox1.Text = "Programming fundamentals" Then
        Selection.TypeText Text:=ComboBox1.Text
        End If
        If ComboBox1.Text = "Data structures and algorithms" Then
        Selection.TypeText Text:=ComboBox1.Text
        End If
        If ComboBox1.Text = "Object Oriented Programming" Then
        Selection.TypeText Text:=ComboBox1.Text
        End If
        If ComboBox1.Text = "Informatics" Then
        Selection.TypeText Text:=ComboBox1.Text
        End If
        If ComboBox1.Text = "Other subject" Then
        Selection.TypeText Text:=TextBox12.Text
        End If
    End If
    Selection.TypeText Text:=Chr(187)
    Dim Taskname As String
    Taskname = TextBox7.Text
    If Len(Taskname) > 0 Or (TextBox7.Visible = False) Then
        If ComboBox2.Text = "Laboratory works" Or ComboBox2.Text = "Practical tasks" Then
        Selection.TypeParagraph
        Selection.Font.Name = "Calibri Light"
        Selection.Font.Size = 16
        Selection.Font.Bold = wdToggle
        Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        Selection.TypeText Text:=Taskname
        Selection.Font.Bold = wdToggle
        End If
        If ComboBox2.Text <> "Laboratory works" Then
        If ComboBox2.Text <> "Practical tasks" Then
        Selection.Font.Name = "Calibri"
        Selection.Font.Size = 12
        End If
        End If
    End If
    Selection.TypeParagraph
    Selection.Font.Name = "Calibri"
    Selection.Font.Size = 12
    Selection.TypeParagraph
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    
    Dim Facultyname, Faculty As String
    
    Facultyname = "Faculty"
    Faculty = "PMI"
    
    If ComboBox2.Text <> "Settlement and graphic tasks" Then
    If ComboBox2.Text <> "Course works" Then
    Selection.Font.Name = "Calibri"
    Selection.Font.Size = 12
    ActiveDocument.Tables.Add Range:=Selection.Range, NumRows:=5, NumColumns:= _
        3, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
        wdAutoFitFixed
    With Selection.Tables(1)
        If .Style <> "Table grid" Then
            .Style = "Table grid"
        End If
        .ApplyStyleHeadingRows = True
        .ApplyStyleFirstColumn = True
        .ApplyStyleRowBands = True

    .Rows.Alignment = wdAlignRowCenter
    .Borders(wdBorderTop).LineStyle = wdLineStyleNone
    .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
    .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
    .Borders(wdBorderRight).LineStyle = wdLineStyleNone
    .Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
    .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
    .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
    .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
    End With
    
    Selection.Rows.HeightRule = wdRowHeightAtLeast
    Selection.Rows.Height = CentimetersToPoints(0.8)
    Selection.Cells.PreferredWidthType = wdPreferredWidthPoints
    
    Selection.MoveRight Unit:=wdCharacter, Count:=3, Extend:=wdExtend
    Selection.MoveDown Unit:=wdLine, Count:=4, Extend:=wdExtend

    Selection.Cells.PreferredWidth = CentimetersToPoints(5.11)
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.MoveDown Unit:=wdLine, Count:=4, Extend:=wdExtend
    Selection.Cells.Merge
    Selection.SelectCell
    Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveDown Unit:=wdLine, Count:=4, Extend:=wdExtend
    Selection.SelectCell
    Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
    Selection.Cells.PreferredWidth = CentimetersToPoints(3.4)
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveDown Unit:=wdLine, Count:=4, Extend:=wdExtend
    Selection.SelectCell
    Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
    Selection.Cells.PreferredWidth = CentimetersToPoints(4.42)
    Selection.MoveLeft Unit:=wdCharacter, Count:=3
    On Error Resume Next
    Selection.InlineShapes.AddPicture FileName:="***", LinkToFile:=False, SaveWithDocument:=True
    On Error GoTo 0
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    With Selection.InlineShapes(1)
    .LockAspectRatio = -1
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
    End With
    Selection.MoveRight Unit:=wdCharacter, Count:=2
    Selection.TypeText Text:=Facultyname & ":"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText Text:=Label2.Caption
    Selection.MoveDown Unit:=wdLine, Count:=1
    If (UBound(Split(TextBox3.Text, vbNewLine)) + 1 > 1) Then
    Selection.TypeText Text:=Label31.Caption
    Else:
    If (ComboBox3.Text = "Male") Then
    Selection.TypeText Text:=Label21.Caption
    End If
    If (ComboBox3.Text = "Female") Then
    Selection.TypeText Text:=Label22.Caption
    End If
    End If
    Selection.MoveDown Unit:=wdLine, Count:=1
    If (UBound(Split(TextBox4.Text, vbNewLine)) + 1 = 1) Then
    Selection.TypeText Text:=Label4.Caption
    Else:
    Selection.TypeText Text:=Label30.Caption
    End If
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText Text:=Label18.Caption
    Selection.MoveUp Unit:=wdLine, Count:=4
    Selection.MoveRight Unit:=wdCharacter, Count:=3
        Selection.TypeText Text:=Faculty
    Selection.MoveDown Unit:=wdLine, Count:=1
        Selection.TypeText Text:=TextBox2.Text
    Selection.MoveDown Unit:=wdLine, Count:=1
        Selection.TypeText Text:=TextBox3.Text
    Selection.MoveDown Unit:=wdLine, Count:=1
        Selection.TypeText Text:=TextBox4.Text
    Selection.MoveDown Unit:=wdLine, Count:=1
        Selection.TypeText Text:=TextBox8.Text
    Selection.MoveUp Unit:=wdLine, Count:=4, Extend:=wdExtend
    Selection.MoveLeft Unit:=wdCharacter, Count:=2, Extend:=wdExtend
    Selection.MoveDown Unit:=wdLine, Count:=1
    End If
    End If
    If ComboBox2.Text = "Settlement and graphic tasks" Then
    ActiveDocument.Tables.Add Range:=Selection.Range, NumRows:=5, NumColumns:= _
        5, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
        wdAutoFitFixed
    With Selection.Tables(1)
        If .Style <> "Table grid" Then
            .Style = "Table grid"
        End If
        .ApplyStyleHeadingRows = True
        .ApplyStyleFirstColumn = True
        .ApplyStyleRowBands = True

    .Borders(wdBorderTop).LineStyle = wdLineStyleNone
    .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
    .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
    .Borders(wdBorderRight).LineStyle = wdLineStyleNone
    .Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
    .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
    .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
    .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
    End With
    
    Selection.Rows.HeightRule = wdRowHeightAtLeast
    Selection.Rows.Height = CentimetersToPoints(0.8)
    Selection.Columns.PreferredWidthType = wdPreferredWidthPoints
    
    Selection.MoveRight Unit:=wdCharacter, Count:=5, Extend:=wdExtend
    Selection.MoveDown Unit:=wdLine, Count:=4, Extend:=wdExtend
    Selection.Font.Name = "Calibri"
    Selection.Font.Size = 12
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Selection.SelectCell
    Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveDown Unit:=wdLine, Count:=4, Extend:=wdExtend
    Selection.SelectCell
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
    Selection.MoveLeft Unit:=wdCharacter, Count:=5
    Selection.MoveDown Unit:=wdLine, Count:=4, Extend:=wdExtend
    Selection.Cells.Merge
    Selection.MoveRight Unit:=wdCharacter, Count:=4, Extend:=wdExtend
    Selection.MoveDown Unit:=wdLine, Count:=4, Extend:=wdExtend
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    On Error Resume Next
    Selection.InlineShapes.AddPicture FileName:="***", LinkToFile:=False, SaveWithDocument:=True
    On Error GoTo 0
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    With Selection.InlineShapes(1)
    .LockAspectRatio = -1
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
    End With
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Selection.Columns.PreferredWidth = CentimetersToPoints(5.11)
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveDown Unit:=wdLine, Count:=4, Extend:=wdExtend
    Selection.Columns.PreferredWidth = CentimetersToPoints(3.4)
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveDown Unit:=wdLine, Count:=4, Extend:=wdExtend
    Selection.Columns.PreferredWidth = CentimetersToPoints(4.42)
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveDown Unit:=wdLine, Count:=4, Extend:=wdExtend
    Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Selection.Columns.PreferredWidth = CentimetersToPoints(2.04)
    Selection.MoveLeft Unit:=wdCharacter, Count:=3
    Selection.TypeText Text:=Facultyname & ":"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText Text:=Label2.Caption
    Selection.MoveDown Unit:=wdLine, Count:=1
    If (UBound(Split(TextBox3.Text, vbNewLine)) + 1 > 1) Then
    Selection.TypeText Text:=Label31.Caption
    Else:
    If (ComboBox3.Text = "Male") Then
    Selection.TypeText Text:=Label21.Caption
    End If
    If (ComboBox3.Text = "Female") Then
    Selection.TypeText Text:=Label22.Caption
    End If
    End If
    Selection.MoveDown Unit:=wdLine, Count:=1
    If (UBound(Split(TextBox4.Text, vbNewLine)) + 1 = 1) Then
    Selection.TypeText Text:=Label4.Caption
    Else:
    Selection.TypeText Text:=Label30.Caption
    End If
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText Text:=Label18.Caption
    Selection.MoveRight Unit:=wdCharacter, Count:=1
        Selection.TypeText Text:=TextBox8.Text
    Selection.MoveUp Unit:=wdLine, Count:=1
        Selection.TypeText Text:=TextBox4.Text
    Selection.MoveUp Unit:=wdLine, Count:=1
        Selection.TypeText Text:=TextBox3.Text
    Selection.MoveUp Unit:=wdLine, Count:=1
        Selection.TypeText Text:=TextBox2.Text
    Selection.MoveUp Unit:=wdLine, Count:=1
        Selection.TypeText Text:=Faculty
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveDown Unit:=wdLine, Count:=2
    Selection.TypeText Text:=Label19.Caption
    Selection.MoveRight Unit:=wdCharacter, Count:=1
        If (TextBox9.Text = 0) Then
        Selection.TypeText Text:=""
        Else
        Selection.TypeText Text:=TextBox9.Text
        End If
        Selection.MoveRight Unit:=wdCharacter, Count:=1
        ActiveWindow.View.ShowXMLMarkup = wdToggle
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:="Signature" & ":"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText Text:="Date" & ":"
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.InsertDateTime DateTimeFormat:="dd.MM.yyyy", InsertAsField:=False _
        , DateLanguage:=wdRussian, CalendarType:=wdCalendarWestern, _
        InsertAsFullWidth:=False
    Selection.MoveRight Unit:=wdCharacter, Count:=32
    Selection.MoveDown Unit:=wdLine, Count:=3
        End If
        If ComboBox2.Text = "Course works" Then
    ActiveDocument.Tables.Add Range:=Selection.Range, NumRows:=5, NumColumns:= _
        5, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
        wdAutoFitFixed
    With Selection.Tables(1)
        If .Style <> "Table grid" Then
            .Style = "Table grid"
        End If
        .ApplyStyleHeadingRows = True
        .ApplyStyleFirstColumn = True
        .ApplyStyleRowBands = True

    .Borders(wdBorderTop).LineStyle = wdLineStyleNone
    .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
    .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
    .Borders(wdBorderRight).LineStyle = wdLineStyleNone
    .Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
    .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
    .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
    .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
    End With
    
    Selection.Rows.HeightRule = wdRowHeightAtLeast
    Selection.Rows.Height = CentimetersToPoints(0.8)
    Selection.Columns.PreferredWidthType = wdPreferredWidthPoints
    
    Selection.MoveRight Unit:=wdCharacter, Count:=5, Extend:=wdExtend
    Selection.MoveDown Unit:=wdLine, Count:=4, Extend:=wdExtend
    Selection.Font.Name = "Calibri"
    Selection.Font.Size = 12
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Selection.SelectCell
    Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveDown Unit:=wdLine, Count:=4, Extend:=wdExtend
    Selection.SelectCell
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
    Selection.MoveLeft Unit:=wdCharacter, Count:=5
    Selection.MoveDown Unit:=wdLine, Count:=4, Extend:=wdExtend
    Selection.Cells.Merge
    Selection.MoveRight Unit:=wdCharacter, Count:=4, Extend:=wdExtend
    Selection.MoveDown Unit:=wdLine, Count:=4, Extend:=wdExtend
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    On Error Resume Next
    Selection.InlineShapes.AddPicture FileName:="***", LinkToFile:=False, SaveWithDocument:=True
    On Error GoTo 0
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    With Selection.InlineShapes(1)
    .LockAspectRatio = -1
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
    End With
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Selection.Columns.PreferredWidth = CentimetersToPoints(5.11)
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveDown Unit:=wdLine, Count:=4, Extend:=wdExtend
    Selection.Columns.PreferredWidth = CentimetersToPoints(3.4)
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveDown Unit:=wdLine, Count:=4, Extend:=wdExtend
    Selection.Columns.PreferredWidth = CentimetersToPoints(4.42)
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveDown Unit:=wdLine, Count:=4, Extend:=wdExtend
    Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Selection.Columns.PreferredWidth = CentimetersToPoints(2.04)
    Selection.MoveLeft Unit:=wdCharacter, Count:=3
    Selection.TypeText Text:=Facultyname & ":"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText Text:=Label2.Caption
    Selection.MoveDown Unit:=wdLine, Count:=1
    If (UBound(Split(TextBox3.Text, vbNewLine)) + 1 > 1) Then
    Selection.TypeText Text:=Label31.Caption
    Else:
    If (ComboBox3.Text = "Male") Then
    Selection.TypeText Text:=Label21.Caption
    End If
    If (ComboBox3.Text = "Female") Then
    Selection.TypeText Text:=Label22.Caption
    End If
    End If
    Selection.MoveDown Unit:=wdLine, Count:=1
    If (UBound(Split(TextBox4.Text, vbNewLine)) + 1 = 1) Then
    Selection.TypeText Text:=Label4.Caption
    Else:
    Selection.TypeText Text:=Label30.Caption
    End If
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText Text:=Label18.Caption
    Selection.MoveRight Unit:=wdCharacter, Count:=1
        Selection.TypeText Text:=TextBox8.Text
    Selection.MoveUp Unit:=wdLine, Count:=1
        Selection.TypeText Text:=TextBox4.Text
    Selection.MoveUp Unit:=wdLine, Count:=1
        Selection.TypeText Text:=TextBox3.Text
    Selection.MoveUp Unit:=wdLine, Count:=1
        Selection.TypeText Text:=TextBox2.Text
    Selection.MoveUp Unit:=wdLine, Count:=1
        Selection.TypeText Text:=Faculty
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:="Mark" & ":"
    Selection.MoveRight Unit:=wdCharacter, Count:=1
        Selection.TypeText Text:="[Mark]"
    Selection.MoveLeft Unit:=wdCharacter, Count:=12
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText Text:="ECTS" & ":"
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:="[ECTS]"
    Selection.MoveLeft Unit:=wdCharacter, Count:=11
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText Text:=Label19.Caption
    Selection.MoveRight Unit:=wdCharacter, Count:=1
        If (TextBox9.Text = 0) Then
        Selection.TypeText Text:=""
        Else
        Selection.TypeText Text:=TextBox9.Text
        End If
        Selection.MoveRight Unit:=wdCharacter, Count:=1
        ActiveWindow.View.ShowXMLMarkup = wdToggle
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
        With Selection.Find
    Selection.HomeKey Unit:=wdLine
        .Text = "[Mark]"
        If (TextBox9.Text = 0) Then
        .Replacement.Text = ""
        Else
        If (TextBox9.Text < 50) Then
        .Replacement.Text = "2"
        Else
        If (TextBox9.Text < 73) Then
        .Replacement.Text = "3"
        Else
        If (TextBox9.Text < 87) Then
        .Replacement.Text = "4"
        Else
        .Replacement.Text = "5"
        End If
        End If
        End If
        End If
        .Forward = True
        .Wrap = wdFindAsk
        Selection.Find.Execute Replace:=wdReplaceAll
        End With
        With Selection.Find
    Selection.HomeKey Unit:=wdLine
        .Text = "[ECTS]"
        If (TextBox9.Text < 50) Then
        .Replacement.Text = ""
        Else
        If (TextBox9.Text < 60) Then
        .Replacement.Text = "E"
        Else
        If (TextBox9.Text < 63) Then
        .Replacement.Text = "D-"
        Else
        If (TextBox9.Text < 67) Then
        .Replacement.Text = "D"
        Else
        If (TextBox9.Text < 70) Then
        .Replacement.Text = "D+"
        Else
        If (TextBox9.Text < 73) Then
        .Replacement.Text = "C-"
        Else
        If (TextBox9.Text < 77) Then
        .Replacement.Text = "C"
        Else
        If (TextBox9.Text < 80) Then
        .Replacement.Text = "C+"
        Else
        If (TextBox9.Text < 83) Then
        .Replacement.Text = "B-"
        Else
        If (TextBox9.Text < 87) Then
        .Replacement.Text = "B"
        Else
        If (TextBox9.Text < 91) Then
        .Replacement.Text = "B+"
        Else
        If (TextBox9.Text < 93) Then
        .Replacement.Text = "A-"
        Else
        If (TextBox9.Text < 98) Then
        .Replacement.Text = "A"
        Else
        .Replacement.Text = "A+"
        End If
        End If
        End If
        End If
        End If
        End If
        End If
        End If
        End If
        End If
        End If
        End If
        End If
        .Forward = True
        .Wrap = wdFindAsk
        Selection.Find.Execute Replace:=wdReplaceAll
        End With
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:="Signature" & ":"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText Text:="Date" & ":"
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.InsertDateTime DateTimeFormat:="dd.MM.yyyy", InsertAsField:=False _
        , DateLanguage:=wdRussian, CalendarType:=wdCalendarWestern, _
        InsertAsFullWidth:=False
    Selection.MoveRight Unit:=wdCharacter, Count:=32
    Selection.MoveDown Unit:=wdLine, Count:=4
    End If
    Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    Selection.TypeParagraph
     If ComboBox2.Text <> "Settlement and graphic tasks" Then
     If ComboBox2.Text <> "Course works" Then
    Selection.TypeParagraph
    End If
    End If
    If ComboBox2.Text = "Settlement and graphic tasks" Then
    Selection.TypeParagraph
    Selection.TypeParagraph
    End If
    If ComboBox2.Text = "Course works" Then
    Selection.TypeParagraph
    Selection.TypeParagraph
    End If
    If ComboBox2.Text <> "Settlement and graphic tasks" Then
    If ComboBox2.Text <> "Course works" Then
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    End If
    End If
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.Font.Name = "Calibri Light"
    Selection.Font.Size = 14
    Selection.TypeText Text:=Text7
    Selection.TypeParagraph
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.InsertDateTime DateTimeFormat:="yyyy", InsertAsField:=False _
        , DateLanguage:=wdRussian, CalendarType:=wdCalendarWestern, _
        InsertAsFullWidth:=False
    ActiveWindow.View.ShowXMLMarkup = wdToggle
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeParagraph
    Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    Selection.Font.Bold = wdToggle
    Selection.TypeText Text:="1. The task"
    Selection.TypeParagraph
    Selection.Font.Bold = wdToggle
    Selection.Font.Name = "Calibri"
    Selection.Font.Size = 12
    Selection.TypeText Text:="Task condition text."
    Selection.TypeParagraph
    Selection.Font.Name = "Calibri Light"
    Selection.Font.Size = 14
    Selection.Font.Bold = wdToggle
    Selection.TypeText Text:="2. Solution"
    Selection.TypeParagraph
    Selection.Font.Bold = wdToggle
    Selection.Font.Name = "Calibri"
    Selection.Font.Size = 12
    Selection.TypeText Text:="Task solution text with illustrations."
    Selection.WholeStory
    Selection.LanguageID = wdRussian
    Selection.NoProofing = False
    Application.CheckLanguage = True
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    ActiveDocument.ShowSpellingErrors = True
    ActiveDocument.ShowGrammaticalErrors = True
    ActiveDocument.ActiveWindow.View.ShowHiddenText = False
    ActiveDocument.RemovePersonalInformation = True
    ActiveDocument.UndoClear
    If (ActiveDocument.ComputeStatistics(wdStatisticPages) < 2) Then
    Unload UserForm1
    Application.Quit SaveChanges:=wdDoNotSaveChanges
    End If
    If OptionButton1.Value = True Then
    If (Application.Version >= "14") Then
    ActiveDocument.SaveAs2 FileName:=ThisDocument.Path & "\" & ComboBox1.Text & "_" & Format(Date, "mmss") & ".docx", FileFormat:=wdFormatDocumentDefault, SaveFormsData:=False, AllowSubstitutions:=False, AddToRecentFiles:=False, LockComments:=True
    End If
    If (Application.Version >= "12" And Application.Version < "14") Then
    ActiveDocument.SaveAs FileName:=ThisDocument.Path & "\" & ComboBox1.Text & "_" & Format(Date, "mmss") & ".docx", FileFormat:=wdFormatDocumentDefault, SaveFormsData:=False, AllowSubstitutions:=False, AddToRecentFiles:=False, LockComments:=True
    End If
    MsgBox "Successfully. The document is saved in a file " & ThisDocument.Path & "\" & ComboBox1.Text & "_" & Format(Date, "mmss") & ".docx"
    Else
    If (Application.Version >= "16") Then
    ActiveDocument.ExportAsFixedFormat2 OutputFileName:=ThisDocument.Path & "\" & ComboBox1.Text & "_" & Format(Date, "mmss") & ".pdf", ExportFormat:=wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:=wdExportOptimizeForPrint, IncludeDocProps:=False, CreateBookmarks:=wdExportCreateNoBookmarks, DocStructureTags:=False, OptimizeForImageQuality:=True
    End If
    If (Application.Version >= "12" And Application.Version < "16") Then
    ActiveDocument.ExportAsFixedFormat OutputFileName:=ThisDocument.Path & "\" & ComboBox1.Text & "_" & Format(Date, "mmss") & ".pdf", ExportFormat:=wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:=wdExportOptimizeForPrint, IncludeDocProps:=False, CreateBookmarks:=wdExportCreateNoBookmarks, DocStructureTags:=False
    End If
    MsgBox "Successfully. The document is saved in a file " & ThisDocument.Path & "\" & ComboBox1.Text & "_" & Format(Date, "mmss") & ".pdf"
End If
Unload UserForm1
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Quit SaveChanges:=wdDoNotSaveChanges
End If
End Sub
Private Sub CommandButton2_Click()
Unload UserForm1
End Sub



Private Sub UserForm_Click()

End Sub
