VERSION 5.00
Begin VB.UserControl cmd 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   DefaultCancel   =   -1  'True
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
End
Attribute VB_Name = "cmd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%              <<< AnaSys Softwares >>>              %
'%                                                    %
'%   DESIGNED BUTTON - copyright ©2001 by Vikrant     %
'%                                                    %
'%  this custom control will emulate the most common  %
'%      command buttons that everyone knows.          %
'%                                                    %
'% it took me about two weeks to develop this control %
'%  and at this time i think it's completely bug free %
'%     ALL THE CODE WAS WRITTEN FROM SCRATCH!!!       %
'%                                                    %
'%   ever wanted to add cool buttons to your app???   %
'%          this is the BEST solution!!!              %
'%                                                    %
'%                                                    %
'%     e-mail: vikrant_thakker@yahoo.com              %
'%                                                    %
'%                                                    %
'%                                                    %
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

'######################################################
'#                    UPDTATE LOG                     #
'#  all times are GMT -03:00                          #
'#                                                    #
'#  November 9 - 03:00 am                             #
'#              · first release                       #
'#                                                    #
'#  November 9 - 05:00 pm                             #
'#              · added ShowFocusRect property        #
'#              · added repaint before triggering the #
'#                click event                         #
'#                                                    #
'#  November 9 - 07:20 pm                             #
'#              · fixed the color shifting so it will #
'#                display the correct color and not a #
'#                weird one.                          #
'#              · improved Java button drawing        #
'#              · added custom colors capability      #
'#                now it looks better than ever COOL! #
'#              · improved Flat button drawing        #
'#                                                    #
'# November 13 - 03:40 pm                             #
'#              · fixed the WinXP button colors and   #
'#                styles. Note that as the colors are #
'#                relative to a base, and for this    #
'#                button i made a color work-around,  #
'#                some colors will be un-reachable    #
'#              · added MouseMove event as requested  #
'#                                                    #
'# November 18 - 10:40 am                             #
'#              · translated all the line methods to  #
'#                API calls. It's now faster than     #
'#                ever. It will also decrease the     #
'#                extra size of your exe!!!           #
'#              · improved Win32 button drawing       #
'#              · moved the direct calls to SetPixel  #
'#                to use less inline .hdc calls       #
'#              · fixed KeyDown/KeyUp events so they  #
'#                now act as they should              #
'#                                                    #
'######################################################


Private Declare Function SetPixel Lib "gdi32" Alias "SetPixelV" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Const COLOR_BTNFACE = 15
Private Const COLOR_BTNSHADOW = 16
Private Const COLOR_BTNTEXT = 18
Private Const COLOR_BTNHIGHLIGHT = 20
Private Const COLOR_BTNDKSHADOW = 21
Private Const COLOR_BTNLIGHT = 22

Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Const DT_LEFT = &H0
Private Const DT_CENTERABS = &H65

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Const RGN_DIFF = 4

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Public Enum ButtonTypes
    [Windows 16-bit] = 1    'the old-fashioned Win16 button
    [Windows 32-bit] = 2    'the classic windows button
    [Windows XP] = 3        'the new brand XP button totally owner-drawn
    [Mac] = 4               'i suppose it looks exactly as a Mac button... i took the style from a GetRight skin!!!
    [Java metal] = 5        'there are also other styles but not so different from windows one
    [Netscape 6] = 6        'this is the button displayed in web-pages, it also appears in some java apps
    [Simple Flat] = 7       'the standard flat button seen on toolbars
End Enum

Public Enum ColorTypes
    [Use Windows] = 1
    [Custom] = 2
    [Force Standard] = 3
End Enum

'events
Public Event Click()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)

'variables
Private MyButtonType As ButtonTypes
Private MyColorType As ColorTypes

Private He As Long  'the height of the button
Private Wi As Long  'the width of the button

Private BackC As Long 'back color
Private ForeC As Long 'fore color

Private elTex As String     'current text
Private TextFont As StdFont 'current font

Private rc As RECT, rc2 As RECT, rc3 As RECT
Private rgnNorm As Long

Private LastButton As Byte, LastKeyDown As Byte
Private isEnabled As Boolean
Private hasFocus As Boolean, showFocusR As Boolean

Private cFace As Long, cLight As Long, cHighLight As Long, cShadow As Long, cDarkShadow As Long, cText As Long

Private lastStat As Byte, TE As String 'used to avoid unnecessary repaints

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    Call UserControl_Click
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
Call Redraw(lastStat, True)
End Sub

Private Sub UserControl_Click()
If (LastButton = 1) And (isEnabled = True) Then
    Call Redraw(0, True) 'be sure that the normal status is drawn
    UserControl.Refresh
    RaiseEvent Click
End If
End Sub

Private Sub UserControl_DblClick()
If LastButton = 1 Then
    Call UserControl_MouseDown(1, 1, 1, 1)
End If
End Sub

Private Sub UserControl_GotFocus()
hasFocus = True
Call Redraw(lastStat, True)
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyDown(KeyCode, Shift)

LastKeyDown = KeyCode
If KeyCode = 32 Then 'spacebar pressed
    Call UserControl_MouseDown(1, 1, 1, 1)
ElseIf (KeyCode = 39) Or (KeyCode = 40) Then 'right and down arrows
    SendKeys "{Tab}"
ElseIf (KeyCode = 37) Or (KeyCode = 38) Then 'left and up arrows
    SendKeys "+{Tab}"
End If
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyUp(KeyCode, Shift)

If (KeyCode = 32) And (LastKeyDown = 32) Then 'spacebar pressed
    Call UserControl_MouseUp(1, 1, 1, 1)
    LastButton = 1
    Call UserControl_Click
End If
End Sub

Private Sub UserControl_LostFocus()
hasFocus = False
Call Redraw(lastStat, True)
End Sub

Private Sub UserControl_Initialize()
LastButton = 1
rc2.Left = 2: rc2.Top = 2
Call SetColors
End Sub

Private Sub UserControl_InitProperties()
    isEnabled = True
    showFocusR = True
    Set TextFont = UserControl.Font
    MyButtonType = [Windows 32-bit]
    MyColorType = [Use Windows]
    BackC = GetSysColor(COLOR_BTNFACE)
    ForeC = GetSysColor(COLOR_BTNTEXT)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
LastButton = Button
If Button <> 2 Then Call Redraw(2, False)
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button < 2 Then
    If X < 0 Or Y < 0 Or X > Wi Or Y > He Then
        'we are outside the button
        Call Redraw(0, False)
    Else
        'we are inside the button
        If Button = 1 Then Call Redraw(2, False)
    End If
End If
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 2 Then Call Redraw(0, False)
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'########## BUTTON PROPERTIES ##########
Public Property Get BackColor() As OLE_COLOR
BackColor = BackC
End Property
Public Property Let BackColor(ByVal theCol As OLE_COLOR)
BackC = theCol
Call SetColors
Call Redraw(lastStat, True)
PropertyChanged "BCOL"
End Property

Public Property Get ForeColor() As OLE_COLOR
ForeColor = ForeC
End Property
Public Property Let ForeColor(ByVal theCol As OLE_COLOR)
ForeC = theCol
Call SetColors
Call Redraw(lastStat, True)
PropertyChanged "FCOL"
End Property

Public Property Get ButtonType() As ButtonTypes
ButtonType = MyButtonType
End Property

Public Property Let ButtonType(ByVal newValue As ButtonTypes)
MyButtonType = newValue
Call UserControl_Resize
Call Redraw(0, True)
PropertyChanged "BTYPE"
End Property

Public Property Get Caption() As String
Caption = elTex
End Property

Public Property Let Caption(ByVal newValue As String)
elTex = newValue
Call SetAccessKeys
Call Redraw(0, True)
PropertyChanged "TX"
End Property

Public Property Get Enabled() As Boolean
Enabled = isEnabled
End Property

Public Property Let Enabled(ByVal newValue As Boolean)
isEnabled = newValue
Call Redraw(0, True)
UserControl.Enabled = isEnabled
PropertyChanged "ENAB"
End Property

Public Property Get Font() As Font
Set Font = TextFont
End Property

Public Property Set Font(ByRef newFont As Font)
Set TextFont = newFont
Set UserControl.Font = TextFont
Call Redraw(0, True)
PropertyChanged "FONT"
End Property

'is very common that a windows user uses custom color
'schemes to view his/her desktop, and is also very
'common that this color scheme has weird colors that
'would alter the nice look of my buttons.
'So if you want to force the button to use the windows
'standard colors you may change this property to "Force Standard"

'UPDATE!!!
'you may now use your custom colors to display the button!!!

Public Property Get ColorScheme() As ColorTypes
ColorScheme = MyColorType
End Property

Public Property Let ColorScheme(ByVal newValue As ColorTypes)
MyColorType = newValue
Call SetColors
Call Redraw(0, True)
PropertyChanged "COLTYPE"
End Property

Public Property Get ShowFocusRect() As Boolean
ShowFocusRect = showFocusR
End Property

Public Property Let ShowFocusRect(ByVal newValue As Boolean)
showFocusR = newValue
Call Redraw(lastStat, True)
PropertyChanged "FOCUSR"
End Property


Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

'########## END OF PROPERTIES ##########

Private Sub UserControl_Resize()
    He = UserControl.ScaleHeight
    Wi = UserControl.ScaleWidth
    rc.Bottom = He: rc.Right = Wi
    rc2.Bottom = He: rc2.Right = Wi
    rc3.Left = 4: rc3.Top = 4: rc3.Right = Wi - 4: rc3.Bottom = He - 4
    
    DeleteObject rgnNorm
    Call MakeRegion
    SetWindowRgn UserControl.hWnd, rgnNorm, True
    
    Call Redraw(0, True)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    MyButtonType = PropBag.ReadProperty("BTYPE", 2)
    elTex = PropBag.ReadProperty("TX", "")
    isEnabled = PropBag.ReadProperty("ENAB", True)
    Set TextFont = PropBag.ReadProperty("FONT", UserControl.Font)
    MyColorType = PropBag.ReadProperty("COLTYPE", 1)
    showFocusR = PropBag.ReadProperty("FOCUSR", True)
    BackC = PropBag.ReadProperty("BCOL", GetSysColor(COLOR_BTNFACE))
    ForeC = PropBag.ReadProperty("FCOL", GetSysColor(COLOR_BTNTEXT))

    UserControl.Enabled = isEnabled
    Set UserControl.Font = TextFont
    Call SetColors
    Call SetAccessKeys
    Call Redraw(0, True)

End Sub

Private Sub UserControl_Terminate()
    DeleteObject rgnNorm
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BTYPE", MyButtonType)
    Call PropBag.WriteProperty("TX", elTex)
    Call PropBag.WriteProperty("ENAB", isEnabled)
    Call PropBag.WriteProperty("FONT", TextFont)
    Call PropBag.WriteProperty("COLTYPE", MyColorType)
    Call PropBag.WriteProperty("FOCUSR", showFocusR)
    Call PropBag.WriteProperty("BCOL", BackC)
    Call PropBag.WriteProperty("FCOL", ForeC)
End Sub

Private Sub Redraw(ByVal curStat As Byte, ByVal Force As Boolean)

'here is the CORE of the button, everything is drawn here
'it's not well commented but i think that everything is
'pretty self explanatory...

If Force = False Then 'check drawing redundancy
    If (curStat = lastStat) And (TE = elTex) Then Exit Sub
End If

If He = 0 Then Exit Sub 'we don't want errors

lastStat = curStat
TE = elTex

Dim i As Long, stepXP1 As Single, XPface As Long
Dim preFocusValue As Boolean

preFocusValue = hasFocus 'save this value to restore it later
If hasFocus = True Then hasFocus = ShowFocusRect

With UserControl
.Cls
DrawRectangle 0, 0, Wi, He, cFace

If isEnabled = True Then
    SetTextColor .hdc, cText 'restore font color
    If curStat = 0 Then
'#@#@#@#@#@# BUTTON NORMAL STATE #@#@#@#@#@#
        Select Case MyButtonType
            Case 1 'Windows 16-bit
                DrawText .hdc, elTex, Len(elTex), rc, DT_CENTERABS
                DrawLine 1, 0, Wi - 1, 0, cDarkShadow
                DrawLine 1, He - 1, Wi - 1, He - 1, cDarkShadow
                DrawLine 0, 1, 0, He - 1, cDarkShadow
                DrawLine Wi - 1, 1, Wi - 1, He - 1, cDarkShadow
                DrawRectangle 1, 1, Wi - 2, He - 2, cHighLight, True
                DrawRectangle 2, 2, Wi - 4, He - 4, cHighLight, True
                DrawLine Wi - 2, 1, Wi - 2, He - 1, cShadow
                DrawLine Wi - 3, 2, Wi - 3, He - 1, cShadow
                DrawLine 1, He - 2, Wi - 1, He - 2, cShadow
                DrawLine 2, He - 3, Wi - 2, He - 3, cShadow
                If hasFocus = True Then DrawFocusR
            Case 2 'Windows 32-bit
                DrawText .hdc, elTex, Len(elTex), rc, DT_CENTERABS
                If (Ambient.DisplayAsDefault = True) And (showFocusR = True) Then
                    DrawRectangle 1, 1, Wi - 2, He - 2, cHighLight, True
                    DrawRectangle 2, 2, Wi - 4, He - 4, cLight, True
                    DrawLine Wi - 2, 1, Wi - 2, He - 1, cDarkShadow
                    DrawLine Wi - 3, 2, Wi - 3, He - 1, cShadow
                    DrawLine 1, He - 2, Wi - 1, He - 2, cDarkShadow
                    DrawLine 2, He - 3, Wi - 2, He - 3, cShadow
                    If hasFocus = True Then DrawFocusR
                    DrawRectangle 0, 0, Wi, He, cDarkShadow, True
                Else
                    DrawRectangle 0, 0, Wi - 1, He - 1, cHighLight, True
                    DrawRectangle 1, 1, Wi - 2, He - 2, cLight, True
                    DrawLine Wi - 1, 0, Wi - 1, He, cDarkShadow
                    DrawLine Wi - 2, 1, Wi - 2, He - 1, cShadow
                    DrawLine 0, He - 1, Wi - 1, He - 1, cDarkShadow
                    DrawLine 1, He - 2, Wi - 2, He - 2, cShadow
                End If
            Case 3 'Windows XP
                stepXP1 = 25 / He
                XPface = ShiftColor(cFace, &H30, True)
                For i = 1 To He
                    DrawLine 0, i, Wi, i, ShiftColor(XPface, -stepXP1 * i, True)
                Next
                SetTextColor UserControl.hdc, cText
                DrawText .hdc, elTex, Len(elTex), rc, DT_CENTERABS
                DrawLine 2, 0, Wi - 2, 0, &H733C00
                DrawLine 2, He - 1, Wi - 2, He - 1, &H733C00
                DrawLine 0, 2, 0, He - 2, &H733C00
                DrawLine Wi - 1, 2, Wi - 1, He - 2, &H733C00
                mSetPixel 1, 1, &H7B4D10
                mSetPixel 1, He - 2, &H7B4D10
                mSetPixel Wi - 2, 1, &H7B4D10
                mSetPixel Wi - 2, He - 2, &H7B4D10
                
                If (hasFocus = True) Or ((Ambient.DisplayAsDefault = True) And (showFocusR = True)) Then
                    DrawRectangle 1, 2, Wi - 2, He - 4, &HE7AE8C, True
                    DrawLine 2, He - 2, Wi - 2, He - 2, &HEF826B
                    DrawLine 2, 1, Wi - 2, 1, &HFFE7CE
                    DrawLine 1, 2, Wi - 1, 2, &HF7D7BD
                    
                    DrawLine 2, 3, 2, He - 3, &HF0D1B5
                    DrawLine Wi - 3, 3, Wi - 3, He - 3, &HF0D1B5
                Else 'we do not draw the bevel always because the above code would repaint over it
                    DrawLine 2, He - 2, Wi - 2, He - 2, ShiftColor(XPface, -&H30, True)
                    DrawLine 1, He - 3, Wi - 2, He - 3, ShiftColor(XPface, -&H20, True)
                    DrawLine Wi - 2, 2, Wi - 2, He - 2, ShiftColor(XPface, -&H24, True)
                    DrawLine Wi - 3, 3, Wi - 3, He - 3, ShiftColor(XPface, -&H18, True)
                    DrawLine 2, 1, Wi - 2, 1, ShiftColor(XPface, &H10, True)
                    DrawLine 1, 2, Wi - 2, 2, ShiftColor(XPface, &HA, True)
                    DrawLine 1, 2, 1, He - 2, ShiftColor(XPface, -&H5, True)
                    DrawLine 2, 3, 2, He - 3, ShiftColor(XPface, -&HA, True)
                End If
            Case 4 'Mac
                DrawRectangle 1, 1, Wi - 2, He - 2, cLight
                DrawText .hdc, elTex, Len(elTex), rc, DT_CENTERABS
                DrawLine 2, 0, Wi - 2, 0, cDarkShadow
                DrawLine 2, He - 1, Wi - 2, He - 1, cDarkShadow
                DrawLine 0, 2, 0, He - 2, cDarkShadow
                DrawLine Wi - 1, 2, Wi - 1, He - 2, cDarkShadow
                mSetPixel 1, 1, cDarkShadow
                mSetPixel 1, He - 2, cDarkShadow
                mSetPixel Wi - 2, 1, cDarkShadow
                mSetPixel Wi - 2, He - 2, cDarkShadow
                mSetPixel 1, 2, cFace
                mSetPixel 2, 1, cFace
                DrawLine 3, 2, Wi - 3, 2, cHighLight
                DrawLine 2, 2, 2, He - 3, cHighLight
                mSetPixel 3, 3, cHighLight
                DrawLine Wi - 3, 1, Wi - 3, He - 3, cFace
                DrawLine 1, He - 3, Wi - 3, He - 3, cFace
                mSetPixel Wi - 4, He - 4, cFace
                DrawLine Wi - 2, 3, Wi - 2, He - 2, cShadow
                DrawLine 3, He - 2, Wi - 2, He - 2, cShadow
                mSetPixel Wi - 3, He - 3, cShadow
                mSetPixel 2, He - 2, cFace
                mSetPixel 2, He - 3, cLight
                mSetPixel Wi - 2, 2, cFace
                mSetPixel Wi - 3, 2, cLight
            Case 5 'Java
                .FontBold = True
                DrawRectangle 1, 1, Wi - 1, He - 1, ShiftColor(cFace, &HC)
                DrawText .hdc, elTex, Len(elTex), rc, DT_CENTERABS
                DrawRectangle 1, 1, Wi - 1, He - 1, cHighLight, True
                DrawRectangle 0, 0, Wi - 1, He - 1, ShiftColor(cShadow, -&H1A), True
                mSetPixel 1, He - 2, ShiftColor(cShadow, &H1A)
                mSetPixel Wi - 2, 1, ShiftColor(cShadow, &H1A)
                If hasFocus = True Then DrawRectangle (Wi - UserControl.TextWidth(elTex)) \ 2 - 3, (He - UserControl.TextHeight(elTex)) \ 2 - 1, UserControl.TextWidth(elTex) + 6, UserControl.TextHeight(elTex) + 2, &HCC9999, True
                .FontBold = TextFont.Bold
            Case 6 'Netscape
                DrawText .hdc, elTex, Len(elTex), rc, DT_CENTERABS
                DrawRectangle 0, 0, Wi, He, ShiftColor(cLight, &H8), True
                DrawRectangle 1, 1, Wi - 2, He - 2, ShiftColor(cLight, &H8), True
                DrawLine Wi - 1, 0, Wi - 1, He, cShadow
                DrawLine Wi - 2, 1, Wi - 2, He - 1, cShadow
                DrawLine 0, He - 1, Wi, He - 1, cShadow
                DrawLine 1, He - 2, Wi - 1, He - 2, cShadow
                If hasFocus = True Then DrawFocusR
             Case 7 'Flat
                DrawText .hdc, elTex, Len(elTex), rc, DT_CENTERABS
                DrawRectangle 0, 0, Wi, He, cHighLight, True
                DrawLine Wi - 1, 0, Wi - 1, He, cShadow
                DrawLine 0, He - 1, Wi, He - 1, cShadow
                If hasFocus = True Then DrawFocusR
        End Select
    ElseIf curStat = 2 Then
'#@#@#@#@#@# BUTTON IS DOWN #@#@#@#@#@#
        Select Case MyButtonType
            Case 1 'Windows 16-bit
                DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTERABS
                DrawLine 1, 0, Wi - 1, 0, cDarkShadow
                DrawLine 1, He - 1, Wi - 1, He - 1, cDarkShadow
                DrawLine 0, 1, 0, He - 1, cDarkShadow
                DrawLine Wi - 1, 1, Wi - 1, He - 1, cDarkShadow
                DrawRectangle 1, 1, Wi - 2, He - 2, cShadow, True
                DrawRectangle 2, 2, Wi - 4, He - 4, cShadow, True
                DrawLine Wi - 2, 1, Wi - 2, He - 1, cHighLight
                DrawLine Wi - 3, 2, Wi - 3, He - 1, cHighLight
                DrawLine 1, He - 2, Wi - 1, He - 2, cHighLight
                DrawLine 2, He - 3, Wi - 2, He - 3, cHighLight
                If hasFocus = True Then DrawFocusR
            Case 2 'Windows 32-bit
                DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTERABS
                
                If showFocusR = True Then
                    DrawRectangle 0, 0, Wi, He, cDarkShadow, True
                    DrawRectangle 1, 1, Wi - 2, He - 2, cShadow, True
                    If hasFocus = True Then DrawFocusR
                Else
                    DrawRectangle 0, 0, Wi - 1, He - 1, cDarkShadow, True
                    DrawRectangle 1, 1, Wi - 2, He - 2, cShadow, True
                    DrawLine Wi - 1, 0, Wi - 1, He, cHighLight
                    DrawLine Wi - 2, 1, Wi - 2, He - 1, cLight
                    DrawLine 0, He - 1, Wi - 1, He - 1, cHighLight
                    DrawLine 1, He - 2, Wi - 2, He - 2, cLight
                End If
            Case 3 'Windows XP
                stepXP1 = 15 / He
                XPface = ShiftColor(cFace, &H30, True)
                XPface = ShiftColor(XPface, -32, True)
                For i = 1 To He
                    DrawLine 0, He - i, Wi, He - i, ShiftColor(XPface, -stepXP1 * i, True)
                Next
                SetTextColor .hdc, cText
                DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTERABS
                DrawLine 2, 0, Wi - 2, 0, &H733C00
                DrawLine 2, He - 1, Wi - 2, He - 1, &H733C00
                DrawLine 0, 2, 0, He - 2, &H733C00
                DrawLine Wi - 1, 2, Wi - 1, He - 2, &H733C00
                mSetPixel 1, 1, &H7B4D10
                mSetPixel 1, He - 2, &H7B4D10
                mSetPixel Wi - 2, 1, &H7B4D10
                mSetPixel Wi - 2, He - 2, &H7B4D10
                
                DrawLine 2, He - 2, Wi - 2, He - 2, ShiftColor(XPface, &H10, True)
                DrawLine 1, He - 3, Wi - 2, He - 3, ShiftColor(XPface, &HA, True)
                DrawLine Wi - 2, 2, Wi - 2, He - 2, ShiftColor(XPface, &H5, True)
                DrawLine Wi - 3, 3, Wi - 3, He - 3, XPface
                DrawLine 2, 1, Wi - 2, 1, ShiftColor(XPface, -&H20, True)
                DrawLine 1, 2, Wi - 2, 2, ShiftColor(XPface, -&H18, True)
                DrawLine 1, 2, 1, He - 2, ShiftColor(XPface, -&H20, True)
                DrawLine 2, 2, 2, He - 2, ShiftColor(XPface, -&H16, True)
                
'                DrawRectangle 1, 2, Wi - 2, He - 4, &H31B2FF, True
'                DrawLine 2, He - 2,Wi - 2, He - 2, &H96E7&
'                DrawLine 2, 1,Wi - 2, 1, &HCEF3FF
'                DrawLine 1, 2,Wi - 1, 2, &H8CDBFF
'                DrawLine 2, 3,2, He - 3, &H6BCBFF
'                DrawLine Wi - 3, 3,Wi - 3, He - 3, &H6BCBFF
            Case 4 'Mac
                DrawRectangle 1, 1, Wi - 2, He - 2, ShiftColor(cShadow, -&H10)
                SetTextColor .hdc, cLight
                DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTERABS
                DrawLine 2, 0, Wi - 2, 0, cDarkShadow
                DrawLine 2, He - 1, Wi - 2, He - 1, cDarkShadow
                DrawLine 0, 2, 0, He - 2, cDarkShadow
                DrawLine Wi - 1, 2, Wi - 1, He - 2, cDarkShadow
                DrawRectangle 1, 1, Wi - 2, He - 2, ShiftColor(cShadow, -&H40), True
                DrawRectangle 2, 2, Wi - 4, He - 4, ShiftColor(cShadow, -&H20), True
                mSetPixel 2, 2, ShiftColor(cShadow, -&H40)
                mSetPixel 3, 3, ShiftColor(cShadow, -&H20)
                mSetPixel 1, 1, cDarkShadow
                mSetPixel 1, He - 2, cDarkShadow
                mSetPixel Wi - 2, 1, cDarkShadow
                mSetPixel Wi - 2, He - 2, cDarkShadow
                DrawLine Wi - 3, 1, Wi - 3, He - 3, cShadow
                DrawLine 1, He - 3, Wi - 2, He - 3, cShadow
                mSetPixel Wi - 4, He - 4, cShadow
                DrawLine Wi - 2, 3, Wi - 2, He - 2, ShiftColor(cShadow, -&H10)
                DrawLine 3, He - 2, Wi - 2, He - 2, ShiftColor(cShadow, -&H10)
                mSetPixel Wi - 2, He - 3, ShiftColor(cShadow, -&H20)
                mSetPixel Wi - 3, He - 2, ShiftColor(cShadow, -&H20)

                mSetPixel 2, He - 2, ShiftColor(cShadow, -&H20)
                mSetPixel 2, He - 3, ShiftColor(cShadow, -&H10)
                mSetPixel 1, He - 3, ShiftColor(cShadow, -&H10)
                mSetPixel Wi - 2, 2, ShiftColor(cShadow, -&H20)
                mSetPixel Wi - 3, 2, ShiftColor(cShadow, -&H10)
                mSetPixel Wi - 3, 1, ShiftColor(cShadow, -&H10)
            Case 5 'Java
                .FontBold = True
                DrawRectangle 1, 1, Wi - 2, He - 2, ShiftColor(cShadow, &H10), False
                DrawRectangle 0, 0, Wi - 1, He - 1, ShiftColor(cShadow, -&H1A), True
                DrawLine Wi - 1, 1, Wi - 1, He, cHighLight
                DrawLine 1, He - 1, Wi - 1, He - 1, cHighLight
                DrawText .hdc, elTex, Len(elTex), rc, DT_CENTERABS
                If hasFocus = True Then DrawRectangle (Wi - UserControl.TextWidth(elTex)) \ 2 - 3, (He - UserControl.TextHeight(elTex)) \ 2 - 1, UserControl.TextWidth(elTex) + 6, UserControl.TextHeight(elTex) + 2, &HCC9999, True
                .FontBold = TextFont.Bold
            Case 6 'Netscape
                DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTERABS
                DrawRectangle 0, 0, Wi, He, cShadow, True
                DrawRectangle 1, 1, Wi - 2, He - 2, cShadow, True
                DrawLine Wi - 1, 0, Wi - 1, He, ShiftColor(cLight, &H8)
                DrawLine Wi - 2, 1, Wi - 2, He - 1, ShiftColor(cLight, &H8)
                DrawLine 0, He - 1, Wi, He - 1, ShiftColor(cLight, &H8)
                DrawLine 1, He - 2, Wi - 1, He - 2, ShiftColor(cLight, &H8)
                If hasFocus = True Then DrawFocusR
             Case 7 'Flat
                DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTERABS
                DrawRectangle 0, 0, Wi, He, cShadow, True
                DrawLine Wi - 1, 0, Wi - 1, He, cHighLight
                DrawLine 0, He - 1, Wi - 1, He - 1, cHighLight
                If hasFocus = True Then DrawFocusR
        End Select
    End If
Else
'#~#~#~#~#~# DISABLED STATUS #~#~#~#~#~#
    Select Case MyButtonType
        Case 1 'Windows 16-bit
            SetTextColor .hdc, cHighLight
            DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTERABS
            SetTextColor .hdc, cShadow
            DrawText .hdc, elTex, Len(elTex), rc, DT_CENTERABS
            DrawLine 1, 0, Wi - 1, 0, cDarkShadow
            DrawLine 1, He - 1, Wi - 1, He - 1, cDarkShadow
            DrawLine 0, 1, 0, He - 1, cDarkShadow
            DrawLine Wi - 1, 1, Wi - 1, He - 1, cDarkShadow
            DrawRectangle 1, 1, Wi - 2, He - 2, cHighLight, True
            DrawRectangle 2, 2, Wi - 4, He - 4, cHighLight, True
            DrawLine Wi - 2, 1, Wi - 2, He - 1, cShadow
            DrawLine Wi - 3, 2, Wi - 3, He - 1, cShadow
            DrawLine 1, He - 2, Wi - 1, He - 2, cShadow
            DrawLine 2, He - 3, Wi - 2, He - 3, cShadow
        Case 2 'Windows 32-bit
            SetTextColor .hdc, cHighLight
            DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTERABS
            SetTextColor .hdc, cShadow
            DrawText .hdc, elTex, Len(elTex), rc, DT_CENTERABS
            DrawRectangle 0, 0, Wi - 1, He - 1, cHighLight, True
            DrawRectangle 1, 1, Wi - 2, He - 2, cLight, True
            DrawLine Wi - 1, 0, Wi - 1, He, cDarkShadow
            DrawLine Wi - 2, 1, Wi - 2, He - 1, cShadow
            DrawLine 0, He - 1, Wi - 1, He - 1, cDarkShadow
            DrawLine 1, He - 2, Wi - 2, He - 2, cShadow
        Case 3 'Windows XP
            XPface = ShiftColor(cFace, &H30, True)
            DrawRectangle 0, 0, Wi, He, ShiftColor(XPface, -&H18, True)
            SetTextColor .hdc, ShiftColor(XPface, -&H68, True)
            DrawText .hdc, elTex, Len(elTex), rc, DT_CENTERABS
            DrawLine 2, 0, Wi - 2, 0, ShiftColor(XPface, -&H54, True)
            DrawLine 2, He - 1, Wi - 2, He - 1, ShiftColor(XPface, -&H54, True)
            DrawLine 0, 2, 0, He - 2, ShiftColor(XPface, -&H54, True)
            DrawLine Wi - 1, 2, Wi - 1, He - 2, ShiftColor(XPface, -&H54, True)
            mSetPixel 1, 1, ShiftColor(XPface, -&H48, True)
            mSetPixel 1, He - 2, ShiftColor(XPface, -&H48, True)
            mSetPixel Wi - 2, 1, ShiftColor(XPface, -&H48, True)
            mSetPixel Wi - 2, He - 2, ShiftColor(XPface, -&H48, True)
        Case 4 'Mac
            DrawRectangle 1, 1, Wi - 2, He - 2, cLight
            SetTextColor .hdc, cHighLight
            DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTERABS
            SetTextColor .hdc, cShadow
            DrawText .hdc, elTex, Len(elTex), rc, DT_CENTERABS
            DrawLine 2, 0, Wi - 2, 0, cDarkShadow
            DrawLine 2, He - 1, Wi - 2, He - 1, cDarkShadow
            DrawLine 0, 2, 0, He - 2, cDarkShadow
            DrawLine Wi - 1, 2, Wi - 1, He - 2, cDarkShadow
            mSetPixel 1, 1, cDarkShadow
            mSetPixel 1, He - 2, cDarkShadow
            mSetPixel Wi - 2, 1, cDarkShadow
            mSetPixel Wi - 2, He - 2, cDarkShadow
            mSetPixel 1, 2, cFace
            mSetPixel 2, 1, cFace
            DrawLine 3, 2, Wi - 3, 2, cHighLight
            DrawLine 2, 2, 2, He - 3, cHighLight
            mSetPixel 3, 3, cHighLight
            DrawLine Wi - 3, 1, Wi - 3, He - 3, cFace
            DrawLine 1, He - 3, Wi - 3, He - 3, cFace
            mSetPixel Wi - 4, He - 4, cFace
            DrawLine Wi - 2, 3, Wi - 2, He - 2, cShadow
            DrawLine 3, He - 2, Wi - 2, He - 2, cShadow
            mSetPixel Wi - 3, He - 3, cShadow
            mSetPixel 2, He - 2, cFace
            mSetPixel 2, He - 3, cLight
            mSetPixel Wi - 2, 2, cFace
            mSetPixel Wi - 3, 2, cLight
        Case 5 'Java
            .FontBold = True
            SetTextColor .hdc, cShadow
            DrawText .hdc, elTex, Len(elTex), rc, DT_CENTERABS
            DrawRectangle 0, 0, Wi, He, cShadow, True
            .FontBold = TextFont.Bold
        Case 6 'Netscape
            SetTextColor .hdc, cShadow
            DrawText .hdc, elTex, Len(elTex), rc, DT_CENTERABS
            DrawRectangle 0, 0, Wi, He, ShiftColor(cLight, &H8), True
            DrawRectangle 1, 1, Wi - 2, He - 2, ShiftColor(cLight, &H8), True
            DrawLine Wi - 1, 0, Wi - 1, He, cShadow
            DrawLine Wi - 2, 1, Wi - 2, He - 1, cShadow
            DrawLine 0, He - 1, Wi, He - 1, cShadow
            DrawLine 1, He - 2, Wi - 1, He - 2, cShadow
        Case 7 'Flat
            SetTextColor .hdc, cHighLight
            DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTERABS
            SetTextColor .hdc, cShadow
            DrawText .hdc, elTex, Len(elTex), rc, DT_CENTERABS
            DrawRectangle 0, 0, Wi, He, cHighLight, True
            DrawLine Wi - 1, 0, Wi - 1, He, cShadow
            DrawLine 0, He - 1, Wi - 1, He - 1, cShadow
    End Select
End If
End With
'restore focus value
hasFocus = preFocusValue

End Sub

Private Sub DrawRectangle(ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Color As Long, Optional OnlyBorder As Boolean = False)
'this is my custom function to draw rectangles and frames
'it's faster and smoother than using the line method

Dim bRect As RECT
Dim hBrush As Long
Dim Ret As Long

bRect.Left = X
bRect.Top = Y
bRect.Right = X + Width
bRect.Bottom = Y + Height

hBrush = CreateSolidBrush(Color)

If OnlyBorder = False Then
    Ret = FillRect(UserControl.hdc, bRect, hBrush)
Else
    Ret = FrameRect(UserControl.hdc, bRect, hBrush)
End If

Ret = DeleteObject(hBrush)
End Sub

Private Sub DrawLine(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal Color As Long)
'a fast way to draw lines
Dim pt As POINTAPI

UserControl.ForeColor = Color
MoveToEx UserControl.hdc, X1, Y1, pt
LineTo UserControl.hdc, X2, Y2

End Sub

Private Sub mSetPixel(ByVal X As Long, ByVal Y As Long, ByVal Color As Long)
    Call SetPixel(UserControl.hdc, X, Y, Color)
End Sub

Private Sub DrawFocusR()
    SetTextColor UserControl.hdc, cText
    DrawFocusRect UserControl.hdc, rc3
End Sub
Private Sub SetColors()
'this function sets the colors taken as a base to build
'all the other colors and styles.

If MyColorType = Custom Then
    cFace = BackC
    cText = ForeC
    cShadow = ShiftColor(cFace, -&H40)
    cLight = ShiftColor(cFace, &H1F)
    cHighLight = ShiftColor(cFace, &H2F) 'it should be 3F but it looks too lighter
    cDarkShadow = ShiftColor(cFace, -&HC0)
ElseIf MyColorType = [Force Standard] Then
    cFace = &HC0C0C0
    cShadow = &H808080
    cLight = &HDFDFDF
    cDarkShadow = &H0
    cHighLight = &HFFFFFF
    cText = &H0
Else
'if MyColorType is 1 or has not been set then use windows colors
    cFace = GetSysColor(COLOR_BTNFACE)
    cShadow = GetSysColor(COLOR_BTNSHADOW)
    cLight = GetSysColor(COLOR_BTNLIGHT)
    cDarkShadow = GetSysColor(COLOR_BTNDKSHADOW)
    cHighLight = GetSysColor(COLOR_BTNHIGHLIGHT)
    cText = GetSysColor(COLOR_BTNTEXT)
End If
End Sub

Private Sub MakeRegion()
'this function creates the regions to "cut" the UserControl
'so it will be transparent in certain areas

Dim rgn1 As Long, rgn2 As Long
    
    DeleteObject rgnNorm
    rgnNorm = CreateRectRgn(0, 0, Wi, He)
    rgn2 = CreateRectRgn(0, 0, 0, 0)
    
Select Case MyButtonType
    Case 1 'Windows 16-bit
        rgn1 = CreateRectRgn(0, 0, 1, 1)
        CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(0, He, 1, He - 1)
        CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(Wi, 0, Wi - 1, 1)
        CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(Wi, He, Wi - 1, He - 1)
        CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
        DeleteObject rgn1
    Case 3, 4 'Windows XP and Mac
        rgn1 = CreateRectRgn(0, 0, 2, 1)
        CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(0, He, 2, He - 1)
        CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(Wi, 0, Wi - 2, 1)
        CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(Wi, He, Wi - 2, He - 1)
        CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(0, 1, 1, 2)
        CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(0, He - 1, 1, He - 2)
        CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(Wi, 1, Wi - 1, 2)
        CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(Wi, He - 1, Wi - 1, He - 2)
        CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
        DeleteObject rgn1
    Case 5 'Java
        rgn1 = CreateRectRgn(0, He, 1, He - 1)
        CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(Wi, 0, Wi - 1, 1)
        CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
        DeleteObject rgn1
End Select

DeleteObject rgn2
End Sub

Private Sub SetAccessKeys()
'this is a TRUE access keys parser
'i hate seeing how other programmers just check for the
'existence of the ampersand regardless of what follows it

Dim ampersandPos As Long

If Len(elTex) > 1 Then
    ampersandPos = InStr(1, elTex, "&", vbTextCompare)
    If (ampersandPos < Len(elTex)) And (ampersandPos > 0) Then
        If Mid(elTex, ampersandPos + 1, 1) <> "&" Then 'if text is sonething like && then no access key should be assigned, so continue searching
            UserControl.AccessKeys = LCase(Mid(elTex, ampersandPos + 1, 1))
        Else 'do only a second pass to find another ampersand character
            ampersandPos = InStr(ampersandPos + 2, elTex, "&", vbTextCompare)
            If Mid(elTex, ampersandPos + 1, 1) <> "&" Then
                UserControl.AccessKeys = LCase(Mid(elTex, ampersandPos + 1, 1))
            Else
                UserControl.AccessKeys = ""
            End If
        End If
    Else
        UserControl.AccessKeys = ""
    End If
Else
    UserControl.AccessKeys = ""
End If
End Sub

Private Function ShiftColor(ByVal Color As Long, ByVal Value As Long, Optional isXP As Boolean = False) As Long
'this function will add or remove a certain color
'quantity and return the result

Dim Red As Long, Blue As Long, Green As Long

If isXP = False Then
    Blue = ((Color \ &H10000) Mod &H100) + Value
Else
    Blue = ((Color \ &H10000) Mod &H100)
    Blue = Blue + ((Blue * Value) \ &HC0)
End If
Green = ((Color \ &H100) Mod &H100) + Value
Red = (Color And &HFF) + Value
    
    'check red
    If Red < 0 Then
        Red = 0
    ElseIf Red > 255 Then
        Red = 255
    End If
    'check green
    If Green < 0 Then
        Green = 0
    ElseIf Green > 255 Then
        Green = 255
    End If
    'check blue
    If Blue < 0 Then
        Blue = 0
    ElseIf Blue > 255 Then
        Blue = 255
    End If

ShiftColor = RGB(Red, Green, Blue)
End Function
