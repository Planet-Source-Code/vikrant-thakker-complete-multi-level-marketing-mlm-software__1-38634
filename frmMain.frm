VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   Caption         =   $"frmMain.frx":0000
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmMain.frx":00AE
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   8295
      Left            =   90
      TabIndex        =   6
      Top             =   -45
      Width           =   11850
      Begin MLM.cmd cmdMemReg 
         Height          =   615
         Left            =   4500
         TabIndex        =   7
         ToolTipText     =   "Click here to Add/Edit/Del  the details of a  Member"
         Top             =   1575
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1085
         BTYPE           =   5
         TX              =   "&Member Registration"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   10083289
         FCOL            =   255
      End
      Begin MLM.cmd cmdPayDues 
         Height          =   615
         Left            =   4500
         TabIndex        =   8
         ToolTipText     =   "Click here for entering the details of Payment made to each member"
         Top             =   2475
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1085
         BTYPE           =   5
         TX              =   "&Pay Dues"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   10083289
         FCOL            =   255
      End
      Begin MLM.cmd cmdDefault 
         Height          =   615
         Left            =   4500
         TabIndex        =   9
         ToolTipText     =   "Click here to enter the text that you want to be displayed by default in the 'Member Registration Form'"
         Top             =   3375
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1085
         BTYPE           =   5
         TX              =   "&Default"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   10083289
         FCOL            =   255
      End
      Begin MLM.cmd cmdReport 
         Height          =   615
         Left            =   4500
         TabIndex        =   10
         ToolTipText     =   "Click here to view different types of Reports"
         Top             =   4275
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1085
         BTYPE           =   5
         TX              =   "&Reports"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   10083289
         FCOL            =   255
      End
      Begin MLM.cmd cmdAbout 
         Height          =   615
         Left            =   4500
         TabIndex        =   11
         ToolTipText     =   "Click to know About the Developer of this software"
         Top             =   5175
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1085
         BTYPE           =   5
         TX              =   "&About"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   10083289
         FCOL            =   255
      End
      Begin MLM.cmd cmdExit 
         Height          =   615
         Left            =   4500
         TabIndex        =   12
         ToolTipText     =   "You know what it means ;-)"
         Top             =   6975
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1085
         BTYPE           =   5
         TX              =   "E&xit"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   10083289
         FCOL            =   255
      End
      Begin MLM.cmd cmdHelp 
         Height          =   615
         Left            =   4500
         TabIndex        =   13
         ToolTipText     =   "Click to know About the Developer of this software"
         Top             =   6075
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1085
         BTYPE           =   5
         TX              =   "&Introduction and Business Logic"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   10083289
         FCOL            =   255
      End
   End
   Begin VB.CommandButton cmdExit1 
      BackColor       =   &H000080FF&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4050
      Visible         =   0   'False
      Width           =   3165
   End
   Begin VB.CommandButton cmdAbout1 
      BackColor       =   &H000080FF&
      Caption         =   "&About"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3270
      Visible         =   0   'False
      Width           =   3165
   End
   Begin VB.CommandButton cmdReport1 
      BackColor       =   &H000080FF&
      Caption         =   "&Reports"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Visible         =   0   'False
      Width           =   3165
   End
   Begin VB.CommandButton cmdDefault1 
      BackColor       =   &H000080FF&
      Caption         =   "&Default"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1710
      Visible         =   0   'False
      Width           =   3165
   End
   Begin VB.CommandButton cmdPayDues1 
      BackColor       =   &H000080FF&
      Caption         =   "&Pay Dues"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   930
      Visible         =   0   'False
      Width           =   3165
   End
   Begin VB.CommandButton cmdMemReg1 
      BackColor       =   &H000080FF&
      Caption         =   "Member Registration"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   3165
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAbout_Click()
Call HideDesktop
Call HideTaskBar
Load frmAbout
frmAbout.Show
frmMain.Hide
End Sub

Private Sub cmdDefault_Click()
frmMain.Hide
Load frmDefault
frmDefault.Show
End Sub

Private Sub cmdExit_Click()
On Error GoTo eerr
    Ans = MsgBox("Are you sure You want to quit ?", vbYesNo, "Warning")
    If Ans = vbYes Then
    End
    Else
    Exit Sub
    End If
Exit Sub
eerr:
MsgBox Err.Description, vbOKOnly, "MLM"
End Sub


Private Sub cmdHelp_Click()
'Shell App.Path & "Introduction.DOC", vbMaximizedFocus
   Dim lngReturnNumber As Long
    
    'Launch File
'    lngReturnNumber = ShellExecLaunchFile(txtPathFile.Text, txtArguments.Text, txtStartPath.Text)

lngReturnNumber = ShellExecLaunchFile(App.Path & "\Introduction.Doc", "", "")
    
    'Check for Errors
    If lngReturnNumber < 33 Then
        Call ShellExecLaunchErr(lngReturnNumber, True)
        Exit Sub
    End If
End Sub

Private Sub cmdMemReg_Click()
Load frmReg
frmReg.Show
frmMain.Hide
End Sub

Private Sub cmdPayDues_Click()
Load frmPay
frmPay.Show
frmMain.Hide
End Sub

Private Sub cmdReport_Click()
frmMain.Hide
Load frmReports
frmReports.Show
End Sub
Private Sub Command2_Click()
Load Form1
Form1.Show
End Sub

