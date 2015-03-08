VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SRForm 
   Caption         =   "������ ����� / ������"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   11205
   OleObjectBlob   =   "SRForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SRForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Multireplace Macros
' https://github.com/wyfinger/multireplace
' (C) Wyfinger / wyfinger@mail.ru / 2015

Option Explicit
'
' ��� ���������� ����� ��� ���������� ��������� ������� ������ ������ (Case/Matchword)
Dim CB2 As Boolean
Dim CB3 As Boolean

Private Sub ComButInsertSpetsSimvol_Click()
  '
  ' ��������������� ��������� ������
  Dim sel As Long
  If ComboBoxSpetsSimvol.ListIndex = -1 Then Exit Sub
  sel = ReplaceBox.SelStart
  ReplaceBox.Text = Left(ReplaceBox.Text, sel) & Mid(ComboBoxSpetsSimvol.List(ComboBoxSpetsSimvol.ListIndex), 2, 1) & Right(ReplaceBox.Text, Len(ReplaceBox.Text) - sel)
End Sub

Private Sub UserForm_Initialize()
  '
  ' ��������������� ��������� ������
  SearchBox.Text = GetSetting("SRMacros", "Settings", "SearchList", "")
  ReplaceBox.Text = GetSetting("SRMacros", "Settings", "ReplaceList", "")
  CheckBox1 = CBool(GetSetting("SRMacros", "Settings", "UseWildcards", True))
  CheckBox2 = CBool(GetSetting("SRMacros", "Settings", "WholeWord", False))
  CheckBox3 = CBool(GetSetting("SRMacros", "Settings", "MatchCase", False))
  CheckBox4 = CBool(GetSetting("SRMacros", "Settings", "TrackRevisions", False))
  CheckBox5 = CBool(GetSetting("SRMacros", "Settings", "UseRegExp", False))
  Dim i As Integer
  ComboBoxSpetsSimvol.Clear
  For i = 1 To 255
    ComboBoxSpetsSimvol.AddItem "'" & Chr(i) & "' ( #" & i & " )"
  Next
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  '
  ' ��������� ��������� � �������
  SaveSetting "SRMacros", "Settings", "SearchList", SearchBox.Text
  SaveSetting "SRMacros", "Settings", "ReplaceList", ReplaceBox.Text
  SaveSetting "SRMacros", "Settings", "UseWildcards", CheckBox1
  SaveSetting "SRMacros", "Settings", "WholeWord", CheckBox2
  SaveSetting "SRMacros", "Settings", "MatchCase", CheckBox3
  SaveSetting "SRMacros", "Settings", "TrackRevisions", CheckBox4
  SaveSetting "SRMacros", "Settings", "UseRegExp", CheckBox5
End Sub

Private Sub CheckBox1_Click()
  '
  ' ��� ��������� ����� "�������������� �����" ������ ����� � ����� "���������� ���������",
  ' � ����� ������� ����� "����� �������" � "��������� �������" �����������
  If CheckBox1.Value = True Then
    CB2 = CheckBox2
    CB3 = CheckBox3
    CheckBox2 = True
    CheckBox3 = True
    CheckBox2.Enabled = False
    CheckBox3.Enabled = False
    CheckBox5 = False
  Else
    CheckBox2 = CB2
    CheckBox3 = CB3
    CheckBox2.Enabled = True
    CheckBox3.Enabled = True
  End If
End Sub

Private Sub CheckBox5_Click()
  '
  ' ��� ��������� ����� "���������� ���������" ������ ����� "�������������� �����" �,
  ' ������������� ������� ���������� ����� "����� �������"
  CheckBox1 = Not CheckBox5
  CheckBox2.Enabled = Not CheckBox5
End Sub
 
Private Function SelectRegExpEntry(RegExp, Selct)
  '
  ' ���� ����� �� regexp ��������� ������� - ������ ��������� � ��������� ���,
  ' �������� � ��������� �������, �.�. ������������ ��� ������ �� ���� ��������� �
  ' �� ������������
  Dim reSearch
  ActiveDocument.Range(0, 0).Select
  If RegExp.Count > 0 Then
    Set reSearch = Selct
    With reSearch.Find
      .Text = RegExp(0).Value
      .MatchWildcards = False
      .MatchWholeWord = False
      .MatchCase = False
      .Forward = True
      .Wrap = wdFindContinue
      .Execute
    End With
    If reSearch.Find.Found = True Then
      reSearch.Select
      SelectRegExpEntry = True
    End If
  End If
  Set reSearch = Nothing
End Function

Private Sub CommandButton1_Click()
  '
  ' ���� ������� �������������� ������ ����� � ��������� ������� ���������� ���������
  ' ������� ������ (���������)
  ' ��������� ����� ��� ������ � ������
  Dim Words
  Words = SRForm.SearchBox.Text
  Words = Split(Words, Chr(13) & Chr(10))
  ' ���� �� ���� ������ ������ � ���� � ����������� �� ������
  Dim i, w
  For i = LBound(Words) To UBound(Words)
    w = Trim(Words(i))
    ' ���� ����� ��� ������ ������ - ��������� �����
    If w = "" Then Exit Sub
    If CheckBox5 Then                  ' ����� ���������� ����������
      ' ����� � �������������� ����������� ���������
      Dim regFinder, findRes
      Set regFinder = CreateObject("VBScript.RegExp")
      regFinder.Pattern = w
      regFinder.IgnoreCase = Not CheckBox3
      regFinder.Global = True
      Set findRes = regFinder.Execute(ActiveDocument.Content)
      ' ���� ���-�� ���� ������� ������� ���������
      If SelectRegExpEntry(findRes, Selection) Then Exit Sub
      ' ��������� �������� � ������������
      Dim sect, head, foot
      For Each sect In ActiveDocument.Sections
        ' ������
        For Each head In sect.Headers
          Set findRes = regFinder.Execute(head.Range.Text)
          If SelectRegExpEntry(findRes, head.Range) Then Exit Sub
        Next
        ' ������
        For Each foot In sect.Footers
          Set findRes = regFinder.Execute(foot.Range.Text)
          If SelectRegExpEntry(findRes, foot.Range) Then Exit Sub
        Next
      Next
    Else                               ' ����� ���������� ������ Word
      ' �����/������ ���������� ������ Word, � ��������������� ������� ��� ���
      Dim FRez
      Set FRez = ActiveDocument.Content
      With FRez.Find
        .Text = w
        .MatchWildcards = CheckBox1.Value
        .MatchWholeWord = CheckBox2.Value
        .MatchCase = CheckBox3.Value
        .Forward = True
        .Wrap = wdFindStop
        .Execute
      End With
      ' �����, ���� ���-�� ������� - �������
      If FRez.Find.Found = True Then
        FRez.Select
        Exit Sub
      End If
      ' ��������� �������� � ������������
      For Each sect In ActiveDocument.Sections
        ' ������
        For Each head In sect.Headers
          Set FRez = head.Range
          With FRez.Find
            .Text = w
            .MatchWildcards = CheckBox1.Value
            .MatchWholeWord = CheckBox2.Value
            .MatchCase = CheckBox3.Value
            .Forward = True
            .Wrap = wdFindStop
            .Execute
          End With
          If FRez.Find.Found = True Then
            FRez.Select
            Exit Sub
          End If
        Next
        ' ������
        For Each head In sect.Headers
          Set FRez = foot.Range
          With FRez.Find
            .Text = w
            .MatchWildcards = CheckBox1.Value
            .MatchWholeWord = CheckBox2.Value
            .MatchCase = CheckBox3.Value
            .Forward = True
            .Wrap = wdFindStop
            .Execute
          End With
          If FRez.Find.Found = True Then
            FRez.Select
            Exit Sub
          End If
        Next
      Next
    End If
  Next
End Sub

Sub CommandButton3_Click()
  '
  ' ���� ������� �������������� ������ �� ����� ���������, �����
  ' ������ ���������� ������������ ��� �������� ��������� ������
  ' ��������� ����� ��� ������ � ������ � ������
  Dim Words, Repl
  Words = SRForm.SearchBox.Text
  Words = Split(Words, Chr(13) & Chr(10))
  Repl = SRForm.ReplaceBox.Text
  Repl = Split(Repl, Chr(13) & Chr(10))
  ' ������ ��� ������ �������� ��������� �������� �����������, ����� �������
  ' �������� ����� �� ��������� ��������� (���� ����� "�������� ������ ���������" �����������)
  Dim TR
  TR = ActiveDocument.TrackRevisions
  If SRForm.CheckBox4 = True Then ActiveDocument.TrackRevisions = True
  ' ���� �� ���� ������ ������ � ���� � ����������� �� ������
  Dim i, j, w, r
  For i = LBound(Words) To UBound(Words)
    w = Trim(Words(i))
    r = Trim(Repl(i))
    ' ���� ����� ��� ������ ������ - ��������� �����
    If w = "" Then Exit Sub
    If CheckBox5 Then
      ' ����� � �������������� ����������� ���������
      Dim regFinder, regRepl, findRes
      Set regFinder = CreateObject("VBScript.RegExp")
      Set regRepl = CreateObject("VBScript.RegExp")
      regFinder.Pattern = w
      regFinder.IgnoreCase = Not CheckBox3
      regFinder.Global = True
      regRepl.Pattern = w
      regRepl.IgnoreCase = Not CheckBox3
      regRepl.Global = True
      Set findRes = regFinder.Execute(ActiveDocument.Content)
      ' ��������� ������ word ������� ������ ���������� ActiveDocument.Content, ���
      ' �� ��������� �������� ��������� �������� �� �������� ��������������� �����.
      ' � ����������� findRes ������ ��������� ��������� ��������� � ���������,
      ' ��������� findRes(�).Value ������ ����� ������������ ���������� word c
      ' ����� ��������� (������ ����������������� �����), ������� � ������� ������
      Dim reSearch
      ActiveDocument.Range(0, 0).Select
      For j = findRes.Count - 1 To 0 Step -1
        ' ����� � ����� ���������
        Set reSearch = Selection
        With reSearch.Find
          .Text = findRes(j).Value
          .MatchWildcards = False
          .MatchWholeWord = False
          .MatchCase = False
          .Forward = False
          .Wrap = wdFindContinue
          .Execute
        End With
        If reSearch.Find.Found = True Then reSearch.Select
        ' ������ ��� � ��� ���������� ������, ����� ������������ ����� ����������
        Selection.Text = regRepl.Replace(Selection.Text, r)
        ' ��� ���������� ������ � ���� ����������� �����
        Selection.MoveLeft wdWord, 1
      Next
      ' ����� � ������������
      Dim sect, head, foot
      For Each sect In ActiveDocument.Sections
        ' ������
        For Each head In sect.Headers
          Set findRes = regFinder.Execute(head.Range.Text)
          ActiveDocument.Range(0, 0).Select
          For j = findRes.Count - 1 To 0 Step -1
            ' ����� � ����� ���������
            Set reSearch = head.Range
            With reSearch.Find
              .Text = findRes(j).Value
              .MatchWildcards = False
              .MatchWholeWord = False
              .MatchCase = False
              .Forward = False
              .Wrap = wdFindContinue
              .Execute
            End With
            If reSearch.Find.Found = True Then reSearch.Select
            ' ������ ��� � ��� ���������� ������, ����� ������������ ����� ����������
            Selection.Text = regRepl.Replace(Selection.Text, r)
            ' ��� ���������� ������ � ���� ����������� �����
            Selection.MoveLeft wdWord, 1
          Next
        Next
        ' ������
        For Each head In sect.Footers
          Set findRes = regFinder.Execute(head.Range.Text)
          ActiveDocument.Range(0, 0).Select
          For j = findRes.Count - 1 To 0 Step -1
            ' ����� � ����� ���������
            Set reSearch = head.Range
            With reSearch.Find
              .Text = findRes(j).Value
              .MatchWildcards = False
              .MatchWholeWord = False
              .MatchCase = False
              .Forward = False
              .Wrap = wdFindContinue
              .Execute
            End With
            If reSearch.Find.Found = True Then reSearch.Select
            ' ������ ��� � ��� ���������� ������, ����� ������������ ����� ����������
            Selection.Text = regRepl.Replace(Selection.Text, r)
            ' ��� ���������� ������ � ���� ����������� �����
            Selection.MoveLeft wdWord, 1
          Next
        Next
      Next
    Else
      ' �������� ����� / ������ ���������� ������ Word
      Dim FRez
      Set FRez = ActiveDocument.Content
      With FRez.Find
        .Text = w
        .Replacement.Text = r
        .MatchWildcards = CheckBox1.Value
        .MatchWholeWord = CheckBox2.Value
        .MatchCase = CheckBox3.Value
        .Forward = True
        .Execute Replace:=wdReplaceAll
      End With
      ' ��������� �������� � ������������
      For Each sect In ActiveDocument.Sections
        ' ������
        For Each head In sect.Headers
          Set FRez = head.Range
          With FRez.Find
            .Text = w
            .Replacement.Text = r
            .MatchWildcards = CheckBox1.Value
            .MatchWholeWord = CheckBox2.Value
            .MatchCase = CheckBox3.Value
            .Forward = True
            .Execute Replace:=wdReplaceAll
          End With
        Next
        ' ������
        For Each foot In sect.Footers
          Set FRez = foot.Range
          With FRez.Find
            .Text = w
            .Replacement.Text = r
            .MatchWildcards = CheckBox1.Value
            .MatchWholeWord = CheckBox2.Value
            .MatchCase = CheckBox3.Value
            .Forward = True
            .Execute Replace:=wdReplaceAll
          End With
        Next
      Next
    End If
  Next
  ' �������������� ������ ������ ��������� � ���������
  If SRForm.CheckBox4 = True Then ActiveDocument.TrackRevisions = TR
End Sub

Private Sub CommandButton4_Click()
  '
  ' ������� ����� �������
  UserForm_QueryClose 0, 0
  SRForm.Hide
End Sub

Private Sub CommandButton5_Click()
  '
  ' ������� ����� �� ������� ������
  SRFiles.TextBox1.Text = FilesList
  SRFiles.Left = Application.Left + Application.Width / 2
  SRFiles.Show
End Sub

Private Sub CommandButton6_Click()
  '
  ' �������� ��������� ������ ������
  UserForm_QueryClose 0, 0
  SRUnit.ProcessFiles
End Sub
