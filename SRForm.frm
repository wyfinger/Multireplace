VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SRForm 
   Caption         =   "Мульти поиск / замена"
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
' эти переменные нужны для сохранения состояния флажков режима поиска (Case/Matchword)
Dim CB2 As Boolean
Dim CB3 As Boolean

Private Sub ComButInsertSpetsSimvol_Click()
  '
  ' восстанавливаем настройки поиска
  Dim sel As Long
  If ComboBoxSpetsSimvol.ListIndex = -1 Then Exit Sub
  sel = ReplaceBox.SelStart
  ReplaceBox.Text = Left(ReplaceBox.Text, sel) & Mid(ComboBoxSpetsSimvol.List(ComboBoxSpetsSimvol.ListIndex), 2, 1) & Right(ReplaceBox.Text, Len(ReplaceBox.Text) - sel)
End Sub

Private Sub UserForm_Initialize()
  '
  ' восстанавливаем настройки поиска
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
  ' сохраняем настройки в реестре
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
  ' при установке галки "Подстановочные знаки" снимем галку с галки "Регулярные выражения",
  ' а также сделаем галки "Слово целиком" и "Учитывать регистр" неактивными
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
  ' при установке галки "Регулярные выражения" снимем галку "Подстановочные знаки" и,
  ' дополнительно сделаем неактивной галку "Слово целиком"
  CheckBox1 = Not CheckBox5
  CheckBox2.Enabled = Not CheckBox5
End Sub
 
Private Function SelectRegExpEntry(RegExp, Selct)
  '
  ' если поиск по regexp увенчался успехом - найдем вхождение и подсветим его,
  ' вынесено в отдельную функцию, т.к. используется при поиске по телу документа и
  ' по колонтитулам
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
  ' этой кнопкой осуществляется только поиск и выделение первого найденного вхождения
  ' искомой строки (выражения)
  ' прочитать слова для поиска в массив
  Dim Words
  Words = SRForm.SearchBox.Text
  Words = Split(Words, Chr(13) & Chr(10))
  ' идем по всем словам подряд и ищем в зависимости от режима
  Dim i, w
  For i = LBound(Words) To UBound(Words)
    w = Trim(Words(i))
    ' если слово для поиска пустое - прекратим поиск
    If w = "" Then Exit Sub
    If CheckBox5 Then                  ' поиск регулярным выражением
      ' поиск с использованием регулярного выражения
      Dim regFinder, findRes
      Set regFinder = CreateObject("VBScript.RegExp")
      regFinder.Pattern = w
      regFinder.IgnoreCase = Not CheckBox3
      regFinder.Global = True
      Set findRes = regFinder.Execute(ActiveDocument.Content)
      ' если что-то было найдено выделим вхождение
      If SelectRegExpEntry(findRes, Selection) Then Exit Sub
      ' попробуем поискать в колонтитулах
      Dim sect, head, foot
      For Each sect In ActiveDocument.Sections
        ' хедеры
        For Each head In sect.Headers
          Set findRes = regFinder.Execute(head.Range.Text)
          If SelectRegExpEntry(findRes, head.Range) Then Exit Sub
        Next
        ' футеры
        For Each foot In sect.Footers
          Set findRes = regFinder.Execute(foot.Range.Text)
          If SelectRegExpEntry(findRes, foot.Range) Then Exit Sub
        Next
      Next
    Else                               ' поиск средствами самого Word
      ' поиск/замена средствами самого Word, с подстановочными знаками или нет
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
      ' снова, если что-то нашлось - выделим
      If FRez.Find.Found = True Then
        FRez.Select
        Exit Sub
      End If
      ' попробуем поискать в колонтитулах
      For Each sect In ActiveDocument.Sections
        ' хедеры
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
        ' футеры
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
  ' этой кнопкой осуществляется замена по всему документу, также
  ' данный обработчик используется для пакетной обработки файлов
  ' прочитать слова для поиска и замены в массив
  Dim Words, Repl
  Words = SRForm.SearchBox.Text
  Words = Split(Words, Chr(13) & Chr(10))
  Repl = SRForm.ReplaceBox.Text
  Repl = Split(Repl, Chr(13) & Chr(10))
  ' прежде чем искать запомним состояние контроля исправлений, чтобы вернуть
  ' исходный режим по окончании обработки (если галка "Включать запись изменений" установлена)
  Dim TR
  TR = ActiveDocument.TrackRevisions
  If SRForm.CheckBox4 = True Then ActiveDocument.TrackRevisions = True
  ' идем по всем словам подряд и ищем в зависимости от режима
  Dim i, j, w, r
  For i = LBound(Words) To UBound(Words)
    w = Trim(Words(i))
    r = Trim(Repl(i))
    ' если слово для поиска пустое - прекратим поиск
    If w = "" Then Exit Sub
    If CheckBox5 Then
      ' поиск с использованием регулярного выражения
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
      ' поскольку глупый word неверно отдает содержимое ActiveDocument.Content, что
      ' не позволяет выделить найденный фрагмент мы применим хитровывернутый прием.
      ' в результатах findRes найдем последнее найденное вхождение в документе,
      ' используя findRes(Х).Value начнем поиск стандартными средствами word c
      ' конца документа (прямой регистрозависимый поиск), выделим и сделаем замену
      Dim reSearch
      ActiveDocument.Range(0, 0).Select
      For j = findRes.Count - 1 To 0 Step -1
        ' поиск с конца документа
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
        ' теперь уже в нем произведем замену, чтобы подвыражения верно вставились
        Selection.Text = regRepl.Replace(Selection.Text, r)
        ' еще передвинем курсор к краю замененного слова
        Selection.MoveLeft wdWord, 1
      Next
      ' поиск в колонтитулах
      Dim sect, head, foot
      For Each sect In ActiveDocument.Sections
        ' хедеры
        For Each head In sect.Headers
          Set findRes = regFinder.Execute(head.Range.Text)
          ActiveDocument.Range(0, 0).Select
          For j = findRes.Count - 1 To 0 Step -1
            ' поиск с конца документа
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
            ' теперь уже в нем произведем замену, чтобы подвыражения верно вставились
            Selection.Text = regRepl.Replace(Selection.Text, r)
            ' еще передвинем курсор к краю замененного слова
            Selection.MoveLeft wdWord, 1
          Next
        Next
        ' футеры
        For Each head In sect.Footers
          Set findRes = regFinder.Execute(head.Range.Text)
          ActiveDocument.Range(0, 0).Select
          For j = findRes.Count - 1 To 0 Step -1
            ' поиск с конца документа
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
            ' теперь уже в нем произведем замену, чтобы подвыражения верно вставились
            Selection.Text = regRepl.Replace(Selection.Text, r)
            ' еще передвинем курсор к краю замененного слова
            Selection.MoveLeft wdWord, 1
          Next
        Next
      Next
    Else
      ' выполним поиск / замену средствами самого Word
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
      ' попробуем поискать в колонтитулах
      For Each sect In ActiveDocument.Sections
        ' хедеры
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
        ' футеры
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
  ' восстановление режима записи изменений в документе
  If SRForm.CheckBox4 = True Then ActiveDocument.TrackRevisions = TR
End Sub

Private Sub CommandButton4_Click()
  '
  ' закрыть форму макроса
  UserForm_QueryClose 0, 0
  SRForm.Hide
End Sub

Private Sub CommandButton5_Click()
  '
  ' открыть форму со списком файлов
  SRFiles.TextBox1.Text = FilesList
  SRFiles.Left = Application.Left + Application.Width / 2
  SRFiles.Show
End Sub

Private Sub CommandButton6_Click()
  '
  ' пакетная обработка списка файлов
  UserForm_QueryClose 0, 0
  SRUnit.ProcessFiles
End Sub
