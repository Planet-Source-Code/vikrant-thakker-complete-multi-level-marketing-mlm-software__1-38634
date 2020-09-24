VERSION 5.00
Begin VB.Form frmDefault 
   Caption         =   "Default Entry for Member Registration Form"
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   7680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1140
      Left            =   -45
      TabIndex        =   20
      Top             =   6345
      Width           =   11850
      Begin MLM.cmd cmdHelp 
         Height          =   465
         Left            =   3375
         TabIndex        =   9
         Top             =   360
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   820
         BTYPE           =   5
         TX              =   "&Help"
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
         BCOL            =   6871953
         FCOL            =   0
      End
      Begin MLM.cmd cmdExit 
         Height          =   465
         Left            =   4590
         TabIndex        =   10
         Top             =   360
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   820
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
         BCOL            =   6871953
         FCOL            =   0
      End
      Begin MLM.cmd cmdSave 
         Height          =   465
         Left            =   2160
         TabIndex        =   8
         Top             =   360
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   820
         BTYPE           =   5
         TX              =   "&Save"
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
         BCOL            =   6871953
         FCOL            =   0
      End
   End
   Begin VB.CommandButton cmdExit1 
      BackColor       =   &H0041E9D8&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5760
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.CommandButton cmdSave1 
      BackColor       =   &H0041E9D8&
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1215
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5760
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.TextBox txtMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00E6AC7D&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2115
      MaxLength       =   15
      TabIndex        =   7
      ToolTipText     =   $"frmDefault.frx":0000
      Top             =   5220
      Width           =   3555
   End
   Begin VB.TextBox txtExpiry 
      Appearance      =   0  'Flat
      BackColor       =   &H00E6AC7D&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2115
      MaxLength       =   1
      TabIndex        =   6
      Text            =   "3"
      ToolTipText     =   "Enter the no. of months, the membership will be valid."
      Top             =   4500
      Width           =   3555
   End
   Begin VB.TextBox txtCity 
      Appearance      =   0  'Flat
      BackColor       =   &H00E6AC7D&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2100
      MaxLength       =   20
      TabIndex        =   2
      Text            =   "Ahmedabad"
      ToolTipText     =   "Enter the city name that you would like to keep as the default City Name in the Member Registration Form"
      Top             =   1440
      Width           =   3555
   End
   Begin VB.TextBox txtAmt 
      Appearance      =   0  'Flat
      BackColor       =   &H00E6AC7D&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2100
      MaxLength       =   5
      TabIndex        =   5
      Text            =   "300"
      ToolTipText     =   "Enter the Amount that you would charge from the new Member for his membership"
      Top             =   3720
      Width           =   3555
   End
   Begin VB.TextBox txtCountry 
      Appearance      =   0  'Flat
      BackColor       =   &H00E6AC7D&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2115
      MaxLength       =   20
      TabIndex        =   4
      Text            =   "India"
      ToolTipText     =   "Enter the Country name that you would like to keep as the default City Name in the Member Registration Form"
      Top             =   2940
      Width           =   3555
   End
   Begin VB.TextBox txtState 
      Appearance      =   0  'Flat
      BackColor       =   &H00E6AC7D&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2100
      MaxLength       =   20
      TabIndex        =   3
      Text            =   "Gujarat"
      ToolTipText     =   "Enter the State name that you would like to keep as the default City Name in the Member Registration Form"
      Top             =   2160
      Width           =   3555
   End
   Begin VB.OptionButton optN 
      Caption         =   "&No"
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      ToolTipText     =   "Select this if you want to manually enter date in the Date field of Member Registration Form"
      Top             =   360
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.OptionButton optY 
      Caption         =   "&Yes"
      Height          =   375
      Left            =   3060
      TabIndex        =   0
      ToolTipText     =   "Select this if you want to automatically display the current date in the Date field of Member Registration Form"
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "Main Member"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   17
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Expiry Months"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   16
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "City"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   15
      Top             =   1500
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Amount Paid"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   14
      Top             =   3780
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Country"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   13
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "State"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   12
      Top             =   2220
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Do you want the current Date as the default value for the Date field"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   11
      Top             =   360
      Width           =   2715
   End
End
Attribute VB_Name = "frmDefault"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdHelp_Click()
MsgBox "This form is used to enter the text that we want to be entered very frequently in the Member Registration Form... i.e. I live in India, and I know that most of the members that will register with me will also be from India, so in this case each time a new member is made, I will be required to enter the word 'India' in the country field of Member Registration form. It is really annoying and time consuming to enter the same text most of the time... So to overcome this problem, I have made this 'Default Entry form', now the text that you enter in these fields will be visible by default on the 'Member Registration Form', you can change the text manually if you wish. !", vbOKOnly, "Help"
End Sub

Private Sub cmdSave_Click()
On Error GoTo serr
If rsDef.RecordCount = 0 Then
    rsDef.AddNew
End If

If optY.Value = True Then
    rsDef!defaultdate = "Y"
ElseIf optN.Value = True Then
    rsDef!defaultdate = "N"
End If

rsDef!defaultstate = txtState.Text
rsDef!defaultcountry = txtCountry.Text
rsDef!defaultamt = txtAmt.Text
rsDef!defaultCity = txtCity.Text
rsDef!DefaultExpiry = txtExpiry.Text
rsDef!Main = txtMain.Text
rsDef.Update

Exit Sub
serr:
MsgBox Err.Description, vbOKOnly, "MLM"
End Sub

Private Sub cmdExit_Click()
Unload Me
frmMain.Show
End Sub

Private Sub Form_Load()
On Error GoTo ferr
If rsDef.RecordCount = 0 Then
    optN.Value = True
    txtState.Text = "Gujarat"
    txtCountry.Text = "India"
    txtAmt.Text = "300"
    txtCity.Text = "Ahmedabad"
    txtExpiry.Text = "3"
    
End If

If rsDef.RecordCount = 1 Then
    rsDef.MoveFirst
    If rsDef!defaultdate = "Y" Then
        optY.Value = True
    ElseIf rsDef!defaultdate = "N" Then
        optN.Value = True
    End If
    
    txtState.Text = rsDef!defaultstate
    txtCountry.Text = rsDef!defaultcountry
    txtAmt.Text = rsDef!defaultamt
    txtCity.Text = rsDef!defaultCity
    txtExpiry.Text = rsDef!DefaultExpiry
    txtMain.Text = rsDef!Main
End If

Exit Sub
ferr:
MsgBox Err.Description, vbOKOnly, "MLM"
End Sub



Private Sub optN_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub optY_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtAmt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtCity_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtCountry_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub
Private Sub txtExpiry_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub
Private Sub txtMain_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtState_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If

End Sub
