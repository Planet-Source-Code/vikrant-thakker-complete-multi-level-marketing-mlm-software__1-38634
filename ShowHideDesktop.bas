Attribute VB_Name = "ShowHideDesktop"
'Task: This Will Hide And Unhide Your Desktop Icons For Various Reasons. Maybe You Could Tighten Up Security At Home Or The Office.
'Declarations
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Option Explicit
'To Hide Desktop
Public Sub HideDesktop()
Dim hwnd As Long
hwnd = FindWindowEx(0&, 0&, "Progman", vbNullString)
ShowWindow hwnd, 0
End Sub
'Show Desktop
Public Sub ShowDesktop()
Dim hwnd As Long
hwnd = FindWindowEx(0&, 0&, "Progman", vbNullString)
ShowWindow hwnd, 5
End Sub

