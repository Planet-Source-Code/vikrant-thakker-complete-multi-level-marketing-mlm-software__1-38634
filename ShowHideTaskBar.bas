Attribute VB_Name = "ShowHideTaskBar"
'Hiding and Showing the Taskbar
'Declarations
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Const SWP_HIDEWINDOW = &H80
Const SWP_SHOWWINDOW = &H40
'Code to Show the Taskbar
Public Sub ShowTaskBar()
Dim Thwnd As Long
Thwnd = FindWindow("Shell_traywnd", "")
Call SetWindowPos(Thwnd, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
End Sub
'Code to Hide the Taskbar
Public Sub HideTaskBar()
Dim Thwnd As Long
Thwnd = FindWindow("Shell_traywnd", "")
Call SetWindowPos(Thwnd, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
End Sub

