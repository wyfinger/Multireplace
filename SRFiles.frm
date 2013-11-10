VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SRFiles 
   Caption         =   "Файлы для обработки"
   ClientHeight    =   10575
   ClientLeft      =   45
   ClientTop       =   465
   ClientWidth     =   10425
   OleObjectBlob   =   "SRFiles.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SRFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (C) Wyfinger / wyfinger@mail.ru
Option Explicit

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Sub DragAcceptFiles Lib "shell32.dll" (ByVal hwnd As Long, ByVal fAccept As Long)
Const GWL_WNDPROC = (-4)


Private Sub CommandButton1_Click()

  Dim a() As String, fileCount As Long, i As Long
    a = SRUnit.GetFiles(fileCount)
    If (fileCount = 0) Then
        MsgBox "Нет файлов в буфере обмена"
    Else
      SRFiles.TextBox1.Text = ""
        For i = 0 To fileCount - 1
           TextBox1.Text = TextBox1.Text & a(i) & Chr(13) & Chr(10)
        Next
    End If
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'
' очередная убогость VBA, при закрытии формы списка файлов мы должны сохранить
' этот самый список файлов в глобальной переменной, иначе она пропадет
 FilesList = TextBox1.Text
End Sub


Private Sub UserForm_Activate()
'
' этот код может нестабильно работать под отладчиком
 Dim lnghWnd&
 lnghWnd = FindWindow(vbNullString, SRFiles.Caption)
 p = SetWindowLong(lnghWnd, GWL_WNDPROC, AddressOf SRUnit.WindowProc)
 DragAcceptFiles lnghWnd, True
End Sub
