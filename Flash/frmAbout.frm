VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "swflash.ocx"
Begin VB.Form frmAbout 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4530
   ForeColor       =   &H00E0E0E0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3420
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash2 
      Height          =   1455
      Left            =   300
      TabIndex        =   1
      Top             =   1440
      Width           =   3855
      _cx             =   6800
      _cy             =   2566
      Movie           =   "c:\programming\flashabout\2.swf"
      Src             =   "c:\programming\flashabout\2.swf"
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      Stacking        =   "below"
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   750
      TabIndex        =   0
      Top             =   240
      Width           =   2895
      _cx             =   4199410
      _cy             =   4196024
      Movie           =   "c:\programming\flashabout\1.swf"
      Src             =   "c:\programming\flashabout\1.swf"
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      Stacking        =   "below"
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      X1              =   4200
      X2              =   3810
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00FFFFFF&
      X1              =   240
      X2              =   4200
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FFFFFF&
      X1              =   4200
      X2              =   4200
      Y1              =   720
      Y2              =   3000
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      X1              =   240
      X2              =   240
      Y1              =   720
      Y2              =   3000
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      X1              =   630
      X2              =   240
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   630
      X2              =   3780
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   630
      X2              =   3780
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   3780
      X2              =   3780
      Y1              =   120
      Y2              =   1320
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   630
      X2              =   630
      Y1              =   120
      Y2              =   1320
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Flash1_FSCommand(ByVal command As String, ByVal args As String)
    Select Case command
    
    End Select
End Sub

Private Sub Form_Click()
Unload frmAbout
End Sub

Private Sub Form_Load()
    Flash1.Movie = App.Path & "\About.swf"
    Flash2.Movie = App.Path & "\AnaSys.swf"
End Sub
