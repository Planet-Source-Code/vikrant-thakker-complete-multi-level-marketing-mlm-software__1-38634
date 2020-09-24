VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmPay 
   Caption         =   "Pay the Dues"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5820
   LinkTopic       =   "Form2"
   ScaleHeight     =   5235
   ScaleWidth      =   5820
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1770
      Left            =   0
      TabIndex        =   14
      Top             =   3195
      Width           =   6225
      Begin MLM.cmd cmdHelp 
         Height          =   465
         Left            =   1980
         TabIndex        =   5
         Top             =   675
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
         Left            =   3195
         TabIndex        =   6
         Top             =   675
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
         Left            =   765
         TabIndex        =   4
         Top             =   675
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
   Begin VB.TextBox txtDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00E6AC7D&
      Height          =   345
      Left            =   2115
      MaxLength       =   15
      TabIndex        =   0
      ToolTipText     =   "Enter the Date of Payment made to Member"
      Top             =   600
      Width           =   2130
   End
   Begin VB.CommandButton cmdSave1 
      BackColor       =   &H0041E9D8&
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   405
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.CommandButton cmdExit1 
      BackColor       =   &H0041E9D8&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1125
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.TextBox txtPay 
      Appearance      =   0  'Flat
      BackColor       =   &H00E6AC7D&
      Height          =   345
      Left            =   2130
      MaxLength       =   5
      TabIndex        =   3
      ToolTipText     =   "Enter the Amount of Payment made to Member"
      Top             =   2070
      Width           =   2085
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00E6AC7D&
      Height          =   345
      Left            =   2130
      MaxLength       =   25
      TabIndex        =   2
      ToolTipText     =   "Enter the Name of the Member to whom payment has been made"
      Top             =   1560
      Width           =   2085
   End
   Begin VB.TextBox txtID 
      Appearance      =   0  'Flat
      BackColor       =   &H00E6AC7D&
      Height          =   345
      Left            =   2130
      MaxLength       =   15
      TabIndex        =   1
      ToolTipText     =   "Enter the ID NO of the Member to whom payment has beed made"
      Top             =   1080
      Width           =   2085
   End
   Begin MSDataGridLib.DataGrid dg 
      Height          =   7425
      Left            =   6240
      TabIndex        =   9
      Top             =   0
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   13097
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   13031096
      HeadLines       =   1
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Member List"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
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
      Left            =   750
      TabIndex        =   13
      Top             =   600
      Width           =   1035
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
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
      Height          =   405
      Left            =   780
      TabIndex        =   10
      Top             =   2130
      Width           =   1155
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   780
      TabIndex        =   8
      Top             =   1590
      Width           =   1125
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ID.NO"
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
      Left            =   765
      TabIndex        =   7
      Top             =   1050
      Width           =   1035
   End
End
Attribute VB_Name = "frmPay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsMemList As Recordset

Private Sub cmdExit_Click()
Unload Me
frmMain.Show
End Sub

Private Sub cmdHelp_Click()
'MsgBox "This form is used to enter the text that we want to be entered very frequently in the Member Registration Form... i.e. I live in India, and I know that most of the members that will register with me will also be from India, so in this case each time a new member is made, I will be required to enter the word 'India' in the country field of Member Registration form. It is really annoying and time consuming to enter the same text most of the time... So to overcome this problem, I have made this 'Default Entry form', now the text that you enter in these fields will be visible by default on the 'Member Registration Form', you can change the text manually if you wish. !", vbOKOnly, "Help"
MsgBox "This form is to enter the details of Payment made to the registered Member.... The list on right side shows the Registered Members List and their due amount at the current date... To select any member, take the mouse cursor on that member and DblClick on row. U can also select the user by pressing the Enter Key... The IDNo. and UserName of the selected users will entered in the respective textboxes, and then you need to enter the amount paid by the company. So this amount will be deducted in the 'Balance'...", vbOKOnly, "Help"
End Sub

Private Sub cmdSave_Click()
On Error GoTo serr
If Trim(txtID.Text) = "" Or Trim(txtName.Text) = "" Or Trim(txtPay.Text) = "" Or Trim(txtDate.Text) = "" Then
    MsgBox "One of the requird field is Empty !", vbOKOnly, "MLM"
    Exit Sub
End If

rsPay.AddNew
    rsPay!Date = txtDate.Text
    rsPay!IDNO = txtID.Text
    rsPay!Name = txtName.Text
    rsPay!Amt = txtPay.Text
rsPay.Update

If rsMem.RecordCount > 0 Then
    rsMem.MoveFirst
    Do While rsMem.EOF = False
    If txtID.Text = rsMem!IDNO Then ' Or Trim(txtID.Text) = "" Then
    ' MsgBox "RefId Found !"
    rsMem!Paid = rsMem!Paid + txtPay.Text
    rsMem!Balance = rsMem!Balance - txtPay.Text
    rsMem.Update
    
    rsMemList.Requery
    dg.Refresh
    
                Exit Do
    End If
    rsMem.MoveNext
    If rsMem.EOF = True Then
    MsgBox "Please Check The ID.No !"
    Exit Sub
    End If
    Loop
    End If

Exit Sub
serr:
MsgBox Err.Description, vbOKOnly, "MLM"

End Sub

Private Sub Form_Load()
On Error GoTo ferr

Set rsMemList = New ADODB.Recordset
rsMemList.Open "select IdNo,Name,Balance from MastCust", conn, adOpenStatic, adLockOptimistic

Set dg.DataSource = rsMemList

If rsDef!defaultdate = "Y" Then
    txtDate.Text = Date
Else
txtDate.Text = ""
End If


Exit Sub
ferr:
MsgBox Err.Description, vbOKOnly, "MLM"
End Sub
Private Sub dg_DblClick()
On Error GoTo dgerr

txtID.Text = (dg.Columns(0).Text)
txtName.Text = dg.Columns(1).Text
txtPay.SetFocus
txtPay.Text = ""
'txtRefEmail.Text = dg.Columns(2).Text
Exit Sub
dgerr:
MsgBox Err.Description, vbOKOnly, "MLM"
End Sub

Private Sub dg_KeyPress(KeyAscii As Integer)
On Error GoTo dgkerr
If KeyAscii = 13 Then
txtID.Text = (dg.Columns(0).Text)
txtName.Text = dg.Columns(1).Text
txtPay.SetFocus
txtPay.Text = ""
'txtRefEmail.Text = dg.Columns(2).Text
'Call cmdOK_Click
End If

Exit Sub
dgkerr:
MsgBox Err.Description, vbOKOnly, "MLM"
End Sub
Private Sub txtDate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtPay_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub
