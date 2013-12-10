Attribute VB_Name = "SRUnit"
' Multireplace Macros
' https://github.com/wyfinger/multireplace
' (C) Wyfinger / wyfinger@mail.ru
Option Explicit

Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal uFormat As Long) As Long
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal uFormat As Long) As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal drop_handle As Long, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long
Public Declare PtrSafe Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const CF_HDROP As Long = 15
Private Const WM_DROPFILES = &H233
Global p&

' � ���� ���������� ���������� �������� ������ ������ ��� �������� ���������
Global FilesList
 
 
Public Function GetFiles(ByRef fileCount As Long) As String()
'
' ��� ������� ������ ������ ������, ������������� � ����� ������
    Dim hDrop As Long, i As Long
    Dim aFiles() As String, sFileName As String * 1024

    fileCount = 0

    If Not CBool(IsClipboardFormatAvailable(CF_HDROP)) Then Exit Function
    If Not CBool(OpenClipboard(0&)) Then Exit Function

    hDrop = GetClipboardData(CF_HDROP)
    If Not CBool(hDrop) Then GoTo done

    fileCount = DragQueryFile(hDrop, -1, vbNullString, 0)

    ReDim aFiles(fileCount - 1)
    For i = 0 To fileCount - 1
        DragQueryFile hDrop, i, sFileName, Len(sFileName)
        aFiles(i) = Left$(sFileName, InStr(sFileName, vbNullChar) - 1)
    Next
    GetFiles = aFiles
done:
    CloseClipboard
End Function
 
 
Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'
' ������� ��������� ���������, ����� ������ � ������������ WM_DROPFILES
    If uMsg = WM_DROPFILES Then
        Dim Str_Name As String * &HFF, Lng_File&, i
        Lng_File = DragQueryFile(wParam, True, Str_Name, 0)
        SRFiles.TextBox1.Text = ""
        For i = 0 To Lng_File - 1
            Str_Name = ""
            DragQueryFile ByVal wParam, i, Str_Name, Len(Str_Name)
            SRFiles.TextBox1.Text = SRFiles.TextBox1.Text & Replace$(Str_Name, vbNullChar, vbNullString) & Chr(13) & Chr(10)
        Next
    End If
    WindowProc = CallWindowProc(p, hwnd, uMsg, wParam, lParam)
End Function


Sub ShowSRForm()
'
' ��������� ������� ������ �������
 SRForm.Show
End Sub


Public Sub SRProcess()
'
' ��������� ������ ��� �������� ��� ��������� ���������� ������
 On Error Resume Next
   SRForm.CommandButton3_Click
End Sub


Public Function ProcessFiles() As Long
'
' ������� ��������� ���������
 Dim fs
 fs = FilesList
 fs = Split(fs, Chr(13) & Chr(10))

' ��������� � ������������ ����� �� ������
 Dim i, w, WD, TR
 For i = LBound(fs) To UBound(fs)
   On Error GoTo OnError
     w = fs(i)
     If w = "" Then GoTo OnError
     
    ' ������� �������� (� ������������ ���� Word)
     Set WD = Application.Documents.Open(w)
     
    ' ����� "�������� ������ ���������" ����������� ������� ���� �����
     TR = WD.TrackRevisions
     If SRForm.CheckBox4.Value = True Then WD.TrackRevisions = True
     
    ' ����� ������� ��������� ��� ��������� ���������
     WD.Application.Run "SRProcess"
     
    ' ������ ����� ����� ������ �����������, ����� ���
     If SRForm.CheckBox4.Value = True Then WD.TrackRevisions = TR
     
    ' �������� � ������� ��������
     WD.Save
     WD.Close
     
' ���� ����������� ������ - �������� � ���������� ���������
OnError:
 Next
End Function
