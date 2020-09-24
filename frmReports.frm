VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmReports 
   Caption         =   "Reports ....    Real Money Zone"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "REPORTS"
      ForeColor       =   &H00000000&
      Height          =   8655
      Left            =   0
      TabIndex        =   24
      ToolTipText     =   "Select the type of Report that you want to view"
      Top             =   0
      Width           =   3555
      Begin VB.OptionButton optMPS 
         BackColor       =   &H00E6AC7D&
         Caption         =   "Memberwise Payment Sheet"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   330
         TabIndex        =   39
         Top             =   4380
         Width           =   2925
      End
      Begin VB.OptionButton optDPS 
         BackColor       =   &H00E6AC7D&
         Caption         =   "Datewise Payment Sheet"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   330
         TabIndex        =   38
         Top             =   4890
         Width           =   2925
      End
      Begin VB.OptionButton optTotPay 
         BackColor       =   &H00E6AC7D&
         Caption         =   "Total Payment Query"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   330
         TabIndex        =   37
         Top             =   5400
         Width           =   2925
      End
      Begin VB.OptionButton optMemBal 
         BackColor       =   &H00E6AC7D&
         Caption         =   "Memberwise Account Info"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   330
         TabIndex        =   36
         Top             =   5910
         Width           =   2925
      End
      Begin VB.OptionButton optDaily 
         BackColor       =   &H00E6AC7D&
         Caption         =   "&Daily Sales Report"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   330
         TabIndex        =   35
         Top             =   6420
         Width           =   2925
      End
      Begin VB.OptionButton optMExpiry 
         BackColor       =   &H00E6AC7D&
         Caption         =   "Memberwise Account  Expiry "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   330
         TabIndex        =   34
         Top             =   3870
         Width           =   2925
      End
      Begin VB.OptionButton optExpiry 
         BackColor       =   &H00E6AC7D&
         Caption         =   "Datewise Account  Expiry"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   330
         TabIndex        =   33
         Top             =   3360
         Width           =   2925
      End
      Begin VB.OptionButton optReg 
         BackColor       =   &H00E6AC7D&
         Caption         =   "Member Registration Report"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   330
         TabIndex        =   32
         Top             =   840
         Value           =   -1  'True
         Width           =   2925
      End
      Begin VB.OptionButton optCont 
         BackColor       =   &H00E6AC7D&
         Caption         =   "Member's Contact Info"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   330
         TabIndex        =   31
         Top             =   1860
         Width           =   2925
      End
      Begin VB.OptionButton optDueAmt 
         BackColor       =   &H00E6AC7D&
         Caption         =   "Citywse Member List"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   330
         TabIndex        =   30
         Top             =   1350
         Width           =   2925
      End
      Begin VB.OptionButton optCompTree 
         BackColor       =   &H00E6AC7D&
         Caption         =   "Company Tree View"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   330
         TabIndex        =   29
         Top             =   2880
         Width           =   2925
      End
      Begin VB.OptionButton optIndTree 
         BackColor       =   &H00E6AC7D&
         Caption         =   "Indivdual Tree View"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   330
         TabIndex        =   28
         Top             =   2370
         Width           =   2925
      End
      Begin VB.CommandButton cmdExit 
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
         Left            =   1020
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   7650
         Width           =   1365
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   8655
      Left            =   3180
      TabIndex        =   16
      Top             =   0
      Width           =   8715
      Begin Crystal.CrystalReport Cr 
         Left            =   6660
         Top             =   1860
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.TextBox txtRep 
         Height          =   315
         Left            =   6600
         TabIndex        =   25
         Text            =   "MemReg"
         Top             =   1140
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Frame frmdet 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Date"
         Height          =   3135
         Left            =   1320
         TabIndex        =   18
         Top             =   540
         Width           =   5115
         Begin VB.CommandButton cmdShow 
            BackColor       =   &H0041E9D8&
            Caption         =   "&Show"
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
            Left            =   1740
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   " View the Report"
            Top             =   2340
            Width           =   1365
         End
         Begin VB.OptionButton optToday 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Today"
            Height          =   285
            Left            =   360
            TabIndex        =   9
            ToolTipText     =   "Generate the Reports for today's Date"
            Top             =   1530
            Width           =   1455
         End
         Begin VB.OptionButton optSpec 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Specific Date"
            Height          =   270
            Left            =   360
            TabIndex        =   8
            ToolTipText     =   "Generate a report of a particular Day"
            Top             =   1140
            Width           =   1365
         End
         Begin VB.OptionButton optBet 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Between Dates"
            Height          =   285
            Left            =   360
            TabIndex        =   1
            ToolTipText     =   "Generate the report between particular Dates"
            Top             =   765
            Width           =   1545
         End
         Begin VB.OptionButton optAll 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&All Dates"
            Height          =   330
            Left            =   375
            TabIndex        =   0
            ToolTipText     =   "generate the report from beginning to the current date"
            Top             =   360
            Value           =   -1  'True
            Width           =   1230
         End
         Begin VB.Frame framedate 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   1335
            Left            =   2640
            TabIndex        =   19
            Top             =   420
            Visible         =   0   'False
            Width           =   2175
            Begin VB.ComboBox cmbtDay 
               BackColor       =   &H00C0C0C0&
               ForeColor       =   &H00000000&
               Height          =   315
               ItemData        =   "frmReports.frx":0000
               Left            =   840
               List            =   "frmReports.frx":0061
               Style           =   1  'Simple Combo
               TabIndex        =   5
               Text            =   "01"
               Top             =   840
               Width           =   390
            End
            Begin VB.ComboBox cmbTMonth 
               BackColor       =   &H00C0C0C0&
               ForeColor       =   &H00000000&
               Height          =   315
               ItemData        =   "frmReports.frx":00E1
               Left            =   1200
               List            =   "frmReports.frx":0109
               Style           =   1  'Simple Combo
               TabIndex        =   6
               Text            =   "01"
               Top             =   840
               Width           =   390
            End
            Begin VB.ComboBox cmbTYear 
               BackColor       =   &H00C0C0C0&
               ForeColor       =   &H00000000&
               Height          =   315
               ItemData        =   "frmReports.frx":013D
               Left            =   1560
               List            =   "frmReports.frx":0189
               Style           =   1  'Simple Combo
               TabIndex        =   7
               Text            =   "2001"
               Top             =   840
               Width           =   495
            End
            Begin VB.ComboBox cmbfYear 
               BackColor       =   &H00C0C0C0&
               ForeColor       =   &H00000000&
               Height          =   315
               ItemData        =   "frmReports.frx":021D
               Left            =   1560
               List            =   "frmReports.frx":0269
               Style           =   1  'Simple Combo
               TabIndex        =   4
               Text            =   "2001"
               Top             =   240
               Width           =   495
            End
            Begin VB.ComboBox cmbfMonth 
               BackColor       =   &H00C0C0C0&
               ForeColor       =   &H00000000&
               Height          =   315
               ItemData        =   "frmReports.frx":02FD
               Left            =   1200
               List            =   "frmReports.frx":0325
               Style           =   1  'Simple Combo
               TabIndex        =   3
               Text            =   "01"
               Top             =   240
               Width           =   390
            End
            Begin VB.ComboBox cmbfDay 
               BackColor       =   &H00C0C0C0&
               ForeColor       =   &H00000000&
               Height          =   315
               ItemData        =   "frmReports.frx":0359
               Left            =   840
               List            =   "frmReports.frx":03BA
               Style           =   1  'Simple Combo
               TabIndex        =   2
               Text            =   "01"
               Top             =   240
               Width           =   390
            End
            Begin VB.Label lblfrom 
               BackStyle       =   0  'Transparent
               Caption         =   "FROM"
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   120
               TabIndex        =   21
               Top             =   300
               Width           =   795
            End
            Begin VB.Label lblto 
               BackStyle       =   0  'Transparent
               Caption         =   "TO"
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   120
               TabIndex        =   20
               Top             =   840
               Width           =   795
            End
         End
      End
      Begin VB.Frame frmview 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Quick View"
         Height          =   1995
         Left            =   1320
         TabIndex        =   17
         Top             =   4140
         Width           =   5115
         Begin VB.CommandButton cmdLastMonth 
            BackColor       =   &H0041E9D8&
            Caption         =   "Last Month"
            Height          =   375
            Left            =   300
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "View Last Month's Report"
            Top             =   1185
            Width           =   1095
         End
         Begin VB.CommandButton cmdLastWeek 
            BackColor       =   &H0041E9D8&
            Caption         =   "Last Week"
            Height          =   375
            Left            =   285
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "View Last week's Report"
            Top             =   600
            Width           =   1095
         End
         Begin VB.CommandButton cmd3rd 
            BackColor       =   &H0041E9D8&
            Caption         =   "3rd Quarter"
            Height          =   375
            Left            =   3060
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "View reports for July to Sept."
            Top             =   600
            Width           =   1095
         End
         Begin VB.CommandButton cmd2nd 
            BackColor       =   &H0041E9D8&
            Caption         =   "2nd Quarter"
            Height          =   375
            Left            =   1665
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "View reports for April to June"
            Top             =   1185
            Width           =   1050
         End
         Begin VB.CommandButton cmd1st 
            BackColor       =   &H0041E9D8&
            Caption         =   "1st Quarter"
            Height          =   375
            Left            =   1680
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "View the reports for the 1st quarter (Months January to March)"
            Top             =   600
            Width           =   1050
         End
         Begin VB.CommandButton cmd4th 
            BackColor       =   &H0041E9D8&
            Caption         =   "4th Quarter"
            Height          =   375
            Left            =   3060
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "View reports for Oct. to Dec."
            Top             =   1200
            Width           =   1050
         End
      End
      Begin VB.Label lblTDate 
         Caption         =   "Label8"
         Height          =   375
         Left            =   7140
         TabIndex        =   23
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblfDate 
         Caption         =   "Label8"
         Height          =   375
         Left            =   5700
         TabIndex        =   22
         Top             =   480
         Visible         =   0   'False
         Width           =   1275
      End
   End
End
Attribute VB_Name = "frmReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
Unload Me
frmMain.Show
End Sub

Private Sub cmdShow_Click()
On Error GoTo eshow
Call DateValidate
Call Report
Call SelectionFormula
Cr.Action = 1
Exit Sub
eshow:
MsgBox Err.Description, vbOKOnly, "MLM"
End Sub
Private Sub SelectionFormula()
On Error GoTo serr

If optMPS.Value = True Or optDPS.Value = True Or optTotPay.Value = True Then
    If optToday.Value = True Then
        Cr.SelectionFormula = "{PayList.Date}=CurrentDate"
    ElseIf optSpec.Value = True Then
        Cr.SelectionFormula = "{PayList.Date}=Date " & "(" & (lblfDate.Caption) & ")"
    ElseIf optBet.Value = True Then
        Cr.SelectionFormula = "{PayList.Date}>Date " & "(" & (lblfDate.Caption) & ")" & "And " & "{PayList.Date}<Date " & "(" & (lblTDate.Caption) & ")"
    End If
Exit Sub
End If

If optExpiry.Value = True Or optMExpiry.Value = True Then
If optToday.Value = True Then
    Cr.SelectionFormula = "{MastCust.Expiry}=CurrentDate"
ElseIf optSpec.Value = True Then
    Cr.SelectionFormula = "{MastCust.Expiry}=Date " & "(" & (lblfDate.Caption) & ")"
ElseIf optBet.Value = True Then
    Cr.SelectionFormula = "{MastCust.Expiry}>Date " & "(" & (lblfDate.Caption) & ")" & "And " & "{MastCust.Expiry}<Date " & "(" & (lblTDate.Caption) & ")"
End If

ElseIf optExpiry.Value = False And optExpiry.Value = False Then

If optToday.Value = True Then
    Cr.SelectionFormula = "{MastCust.Date}=CurrentDate"
ElseIf optSpec.Value = True Then
    Cr.SelectionFormula = "{MastCust.Date}=Date " & "(" & (lblfDate.Caption) & ")"
ElseIf optBet.Value = True Then
    Cr.SelectionFormula = "{MastCust.Date}>Date " & "(" & (lblfDate.Caption) & ")" & "And " & "{MastCust.Date}<Date " & "(" & (lblTDate.Caption) & ")"
End If
End If
Exit Sub
serr:
MsgBox Err.Description, vbOKOnly, "MLM"
End Sub
Private Sub DateValidate()
On Error GoTo derr
If (cmbtDay.Text > 31 Or cmbtDay.Text < 1) Or (cmbTMonth.Text > 12 Or cmbTMonth.Text < 1) Or (cmbTYear.Text < 2000 Or cmbTYear.Text > 2099) Then
MsgBox "Invalid Date !"
Exit Sub
End If
If (cmbfDay.Text > 31 Or cmbfDay.Text < 1) Or (cmbfMonth.Text > 12 Or cmbfMonth.Text < 1) Or (cmbfYear.Text < 2000 Or cmbfYear.Text > 2099) Then
MsgBox "Invalid Date !"
Exit Sub
End If

If optBet.Value = True Then
    lblfDate.Caption = cmbfYear.Text + "," + cmbfMonth.Text + "," + cmbfDay.Text
    lblTDate.Caption = cmbTYear.Text + "," + cmbTMonth.Text + "," + cmbtDay.Text
ElseIf optSpec.Value = True Then
        lblfDate.Caption = cmbfYear.Text + "," + cmbfMonth.Text + "," + cmbfDay.Text
End If
Exit Sub
derr:
MsgBox Err.Description, vbOKOnly, "MLM"
End Sub
Private Sub Report()
On Error GoTo rerr
If txtRep.Text = "DailySales" Then
Call DailySales
ElseIf txtRep.Text = "MemReg" Then Call MemReg
'ElseIf txtRep.Text = "DailySales" Then Call DailySales
ElseIf txtRep.Text = "Contact" Then Call Contact
ElseIf txtRep.Text = "IndTree" Then Call IndTree
ElseIf txtRep.Text = "CompTree" Then Call CompTree
ElseIf txtRep.Text = "MemExp" Then Call MemExp
ElseIf txtRep.Text = "MemDue" Then Call MemDue
ElseIf txtRep.Text = "MemBal" Then Call MemBal
ElseIf txtRep.Text = "MExp" Then Call MExp
ElseIf txtRep.Text = "CityWise" Then Call Citywise
ElseIf txtRep.Text = "CompTree" Then Call CompTree
ElseIf txtRep.Text = "MembPaySheet" Then Call MembPaySheet
ElseIf txtRep.Text = "DatewisePaySheet" Then Call DatewisePaySheet
ElseIf txtRep.Text = "TotalPay" Then Call TotalPay
End If
Exit Sub
rerr:
MsgBox Err.Description, vbOKOnly, "MLM"
End Sub
Private Sub optDPS_Click()
frmdet.Visible = True
frmview.Visible = True
txtRep.Text = "DatewisePaySheet"
End Sub

Private Sub DatewisePaySheet()
On Error GoTo errdps
Cr.Reset
Cr.ReportTitle = "Datewise Payment Sheet"
Cr.WindowState = crptMaximized
Cr.WindowTitle = "Datewise Payment Sheet"
Cr.WindowShowGroupTree = True
Cr.DataFiles(0) = App.Path & "\Data\Data97.mdb"
Cr.ReportFileName = App.Path & "\Reports\rptDatewisePaymentSheet.rpt"

Cr.Destination = crptToWindow
Cr.WindowShowRefreshBtn = True
'Cr.Action = 1
Exit Sub
errdps:
MsgBox Err.Description, vbOKOnly, "MLM"
End Sub

Private Sub optDPS_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub optMemBal_Click()
frmdet.Visible = True
frmview.Visible = True
txtRep.Text = "MemBal"
'Call CompTree
End Sub

Private Sub MemBal()
On Error GoTo membalerr
Cr.Reset
Cr.ReportTitle = "Member Account Info"
Cr.WindowState = crptMaximized
Cr.WindowTitle = "Daily Sales"
Cr.WindowShowGroupTree = True
Cr.DataFiles(0) = App.Path & "\Data\Data97.mdb"
Cr.ReportFileName = App.Path & "\Reports\rptMemBal.rpt"

Cr.Destination = crptToWindow
Cr.WindowShowRefreshBtn = True
'Cr.Action = 1
Exit Sub
membalerr:
MsgBox Err.Description, vbOKOnly, "MLM"
End Sub

Private Sub optAll_Click()
framedate.Visible = False
End Sub

Private Sub optBet_Click()
framedate.Visible = True
lblfrom.Visible = True
lblto.Visible = True

lblfrom.Caption = "From"
lblto.Caption = "To"

cmbfDay.Visible = True
cmbfMonth.Visible = True
cmbfYear.Visible = True

cmbtDay.Visible = True
cmbTMonth.Visible = True
cmbTYear.Visible = True
End Sub

Private Sub optCompTree_Click()
frmdet.Visible = False
frmview.Visible = False
txtRep.Text = "CompTree"
Call CompTree
End Sub
Private Sub CompTree()
On Error GoTo errcomptree
Cr.Reset
Cr.ReportTitle = "Company Tree"
Cr.WindowState = crptMaximized
Cr.WindowTitle = "Company Tree"
Cr.WindowShowGroupTree = True

Cr.DataFiles(0) = App.Path & "\Data\Data97.mdb"
Cr.ReportFileName = App.Path & "\Reports\rptCompanyTreeView.rpt"


Cr.Destination = crptToWindow
Cr.WindowShowRefreshBtn = True
Cr.Action = 1
Exit Sub
errcomptree:
MsgBox Err.Description, vbOKOnly, "MLM"

End Sub


Private Sub optCompTree_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub optCont_Click()
frmdet.Visible = True
frmview.Visible = True
txtRep.Text = "Contact"
End Sub
Private Sub Contact()
On Error GoTo contacterr
Cr.Reset
Cr.ReportTitle = "Contact List"
Cr.WindowState = crptMaximized
Cr.WindowTitle = "Contact List"
Cr.WindowShowGroupTree = True
Cr.DataFiles(0) = App.Path & "\Data\Data97.mdb"
Cr.ReportFileName = App.Path & "\Reports\rptMemCont.rpt"
Cr.Destination = crptToWindow
Cr.WindowShowRefreshBtn = True
Exit Sub
contacterr:
MsgBox Err.Description, vbOKOnly, "MLM"
End Sub
Private Sub optCont_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub optDaily_Click()
frmdet.Visible = True
frmview.Visible = True
txtRep.Text = "DailySales"
End Sub
Private Sub DailySales()
On Error GoTo dserr
Cr.Reset
Cr.ReportTitle = "Daily Sales Summary"
Cr.WindowState = crptMaximized
Cr.WindowTitle = "Daily Sales"
Cr.WindowShowGroupTree = True
Cr.DataFiles(0) = App.Path & "\Data\Data97.mdb"
Cr.ReportFileName = App.Path & "\Reports\rptDailySales.rpt"
Cr.Destination = crptToWindow
Cr.WindowShowRefreshBtn = True
Exit Sub
dserr:
MsgBox Err.Description, vbOKOnly, "MLM"
End Sub


Private Sub optDaily_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub optDueAmt_Click()
frmdet.Visible = True
frmview.Visible = True
txtRep.Text = "CityWise"
End Sub

Private Sub Citywise()
On Error GoTo citymemberr
Cr.Reset
Cr.ReportTitle = "Citywise Member List"
Cr.WindowState = crptMaximized
Cr.WindowTitle = "rptCitywiseMemberList"
Cr.WindowShowGroupTree = True
Cr.DataFiles(0) = App.Path & "\Data\Data97.mdb"
Cr.ReportFileName = App.Path & "\Reports\rptCitywiseMemberList.rpt"

'Cr.ReportFileName = App.Path & "\Reports\rptDailySales.rpt"

'Cr.ReportFileName = "C:\Sardar\Reports\rptDailySales.rpt"
Cr.Destination = crptToWindow
Cr.WindowShowRefreshBtn = True
Exit Sub
citymemberr:
MsgBox Err.Description, vbOKOnly, "MLM"
End Sub








Private Sub MemDue()
On Error GoTo memdueerr
Cr.Reset
Cr.ReportTitle = "Member Account Details"
Cr.WindowState = crptMaximized
Cr.WindowTitle = "Member Account Details"
Cr.WindowShowGroupTree = True
Cr.DataFiles(0) = App.Path & "\Data\Data97.mdb"
'Cr.ReportFileName = App.Path & "\Reports\rptDailySales.rpt"

'Cr.ReportFileName = "C:\Sardar\Reports\rptDailySales.rpt"
Cr.Destination = crptToWindow
Cr.WindowShowRefreshBtn = True
Exit Sub
memdueerr:
MsgBox Err.Description, vbOKOnly, "MLM"
End Sub
Private Sub optDueAmt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub optExpiry_Click()
frmdet.Visible = True
frmview.Visible = True
txtRep.Text = "MemExp"
End Sub
Private Sub MemExp()
On Error GoTo errmemexp
Cr.Reset
Cr.ReportTitle = "Membership Expiry Report"
Cr.WindowState = crptMaximized
Cr.WindowTitle = "Membership Expiry Report"
Cr.WindowShowGroupTree = True
Cr.DataFiles(0) = App.Path & "\Data\Data97.mdb"
Cr.ReportFileName = App.Path & "\Reports\rptAccExp.rpt"

Cr.Destination = crptToWindow
Cr.WindowShowRefreshBtn = True

Exit Sub
errmemexp:
MsgBox Err.Description, vbOKOnly, "MLM"
End Sub
Private Sub optExpiry_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub optIndTree_Click()
frmdet.Visible = False
frmview.Visible = False
txtRep.Text = "IndTree"
Call IndTree
End Sub
Private Sub IndTree()
On Error GoTo errindtree
Cr.Reset
Cr.ReportTitle = "Indvidual Membership Tree"
Cr.WindowState = crptMaximized
Cr.WindowTitle = "Individual Membership Tree"
Cr.WindowShowGroupTree = True
Cr.DataFiles(0) = App.Path & "\Data\Data97.mdb"
Cr.ReportFileName = App.Path & "\Reports\rptIndTreeView.rpt"
Cr.Destination = crptToWindow
Cr.WindowShowRefreshBtn = True
Cr.Action = 1
Exit Sub
errindtree:
MsgBox Err.Description, vbOKOnly, "MLM"
End Sub
Private Sub optIndTree_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub optMemBal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub optMExpiry_Click()
frmdet.Visible = True
frmview.Visible = True
txtRep.Text = "MExp"
End Sub
Private Sub MExp()
On Error GoTo mexperr
Cr.Reset
Cr.ReportTitle = "Membership Expiry Report"
Cr.WindowState = crptMaximized
Cr.WindowTitle = "Membership Expiry Report"
Cr.WindowShowGroupTree = True
Cr.DataFiles(0) = App.Path & "\Data\Data97.mdb"
Cr.ReportFileName = App.Path & "\Reports\rptMemAccExp.rpt"

Cr.Destination = crptToWindow
Cr.WindowShowRefreshBtn = True
Exit Sub
mexperr:
MsgBox Err.Description, vbOKOnly, "MLM"
End Sub
Private Sub optMExpiry_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub optMPS_Click()
frmdet.Visible = True
frmview.Visible = True
txtRep.Text = "MembPaySheet"
End Sub
Private Sub MembPaySheet()
On Error GoTo errmempaysheet
Cr.Reset
Cr.ReportTitle = "Memberwise Payment Sheet"
Cr.WindowState = crptMaximized
Cr.WindowTitle = "Memberwise Payment Sheet"
Cr.WindowShowGroupTree = True
Cr.DataFiles(0) = App.Path & "\Data\Data97.mdb"
Cr.ReportFileName = App.Path & "\Reports\rptmembpaysheet.rpt"

'Cr.ReportFileName = App.Path & "\Reports\rptDailySales.rpt"

'Cr.ReportFileName = "C:\Sardar\Reports\rptDailySales.rpt"
Cr.Destination = crptToWindow
Cr.WindowShowRefreshBtn = True
Exit Sub
errmempaysheet:
MsgBox Err.Description, vbOKOnly, "MLM"
End Sub

Private Sub optMPS_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub optReg_Click()
frmdet.Visible = True
frmview.Visible = True
txtRep.Text = "MemReg"
End Sub
Private Sub MemReg()
On Error GoTo memregerr
Cr.Reset
Cr.ReportTitle = "Member Registration Report"
Cr.WindowState = crptMaximized
Cr.WindowTitle = "Member Registration Report"
Cr.WindowShowGroupTree = True

Cr.DataFiles(0) = App.Path & "\Data\Data97.mdb"
Cr.ReportFileName = App.Path & "\Reports\rptMemReg.rpt"
Cr.Destination = crptToWindow
Cr.WindowShowRefreshBtn = True
Exit Sub
memregerr:
MsgBox Err.Description, vbOKOnly, "MLM"
End Sub
Private Sub optReg_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub optSpec_Click()
framedate.Visible = True
lblfrom.Caption = "Date"
lblfrom.Visible = True
lblto.Visible = False
lblTDate.Caption = lblfDate.Caption

cmbfDay.Visible = True
cmbfMonth.Visible = True
cmbfYear.Visible = True

cmbtDay.Visible = False
cmbTMonth.Visible = False
cmbTYear.Visible = False
End Sub

Private Sub optToday_Click()
framedate.Visible = False
End Sub
Private Sub cmbfDay_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
Numbers = KeyAscii
    If ((Numbers < 48 Or Numbers > 57) And Numbers <> 8 And Numbers <> 13) Then
        MsgBox "Only Numeric Values allowed", vbCritical, "CyberMan"
        KeyAscii = 0
    End If
End Sub
Private Sub cmbfMonth_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
Numbers = KeyAscii
    If ((Numbers < 48 Or Numbers > 57) And Numbers <> 8 And Numbers <> 13) Then
        MsgBox "Only Numeric Values allowed", vbCritical, "CyberMan"
        KeyAscii = 0
    End If
End Sub

Private Sub cmbfYear_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
Numbers = KeyAscii
    If ((Numbers < 48 Or Numbers > 57) And Numbers <> 8 And Numbers <> 13) Then
        MsgBox "Only Numeric Values allowed", vbCritical, "CyberMan"
        KeyAscii = 0
    End If
End Sub
Private Sub cmbtDay_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
Numbers = KeyAscii
    If ((Numbers < 48 Or Numbers > 57) And Numbers <> 8 And Numbers <> 13) Then
        MsgBox "Only Numeric Values allowed", vbCritical, "CyberMan"
        KeyAscii = 0
    End If
End Sub
Private Sub cmbTMonth_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
Numbers = KeyAscii
    If ((Numbers < 48 Or Numbers > 57) And Numbers <> 8 And Numbers <> 13) Then
        MsgBox "Only Numeric Values allowed", vbCritical, "CyberMan"
        KeyAscii = 0
    End If
End Sub

Private Sub cmbTYear_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
Numbers = KeyAscii
    If ((Numbers < 48 Or Numbers > 57) And Numbers <> 8 And Numbers <> 13) Then
        MsgBox "Only Numeric Values allowed", vbCritical, "CyberMan"
        KeyAscii = 0
    End If
End Sub
Private Sub cmdLastWeek_Click()
Call Report
If optMPS.Value = True Or optDPS.Value = True Or optTotPay.Value = True Then
Cr.SelectionFormula = "{PayList.Date}=LastFullWeek"
Cr.Action = 1
Exit Sub
End If

If optExpiry.Value = True Or optMExpiry.Value = True Then
Cr.SelectionFormula = "{MastCust.Expiry}=LastFullWeek"
ElseIf optExpiry.Value = False And optMExpiry.Value = False Then
Cr.SelectionFormula = "{MastCust.Date}=LastFullWeek"
End If
Cr.Action = 1
End Sub
Private Sub cmdLastMonth_Click()
Call Report
If optMPS.Value = True Or optDPS.Value = True Or optTotPay.Value = True Then
Cr.SelectionFormula = "{PayList.Date}=LastFullMonth"
Cr.Action = 1
Exit Sub
End If


If optExpiry.Value = True Or optMExpiry.Value = True Then
Cr.SelectionFormula = "{MastCust.Expiry}=LastFullMonth"
ElseIf optExpiry.Value = False And optMExpiry.Value = False Then
Cr.SelectionFormula = "{MastCust.Date}=LastFullMonth"
End If
Cr.Action = 1
End Sub
Private Sub cmd1st_Click()
Call Report
If optMPS.Value = True Or optDPS.Value = True Or optTotPay.Value = True Then
Cr.SelectionFormula = "{PayList.Date}=Calendar1stQtr"
Cr.Action = 1
Exit Sub
End If

If optExpiry.Value = True Or optMExpiry.Value = True Then
Cr.SelectionFormula = "{MastCust.Expiry}=Calendar1stQtr"
ElseIf optExpiry.Value = False And optMExpiry.Value = False Then
Cr.SelectionFormula = "{MastCust.Date}=Calendar1stQtr"
End If
Cr.Action = 1
End Sub
Private Sub cmd2nd_Click()
Call Report
If optMPS.Value = True Or optDPS.Value = True Or optTotPay.Value = True Then
Cr.SelectionFormula = "{PayList.Date}=Calendar2ndQtr"
Cr.Action = 1
Exit Sub
End If

If optExpiry.Value = True Or optMExpiry.Value = True Then
Cr.SelectionFormula = "{MastCust.Expiry}=Calendar2ndQtr"
ElseIf optExpiry.Value = False And optMExpiry.Value = False Then
Cr.SelectionFormula = "{MastCust.Date}=Calendar2ndQtr"
End If
Cr.Action = 1
End Sub
Private Sub cmd3rd_Click()
Call Report
If optMPS.Value = True Or optDPS.Value = True Or optTotPay.Value = True Then
Cr.SelectionFormula = "{PayList.Date}=Calendar3rdQtr"
Cr.Action = 1
Exit Sub
End If

If optExpiry.Value = True Or optMExpiry.Value = True Then
Cr.SelectionFormula = "{MastCust.Date}=Calendar3rdQtr"
ElseIf optExpiry.Value = False And optMExpiry.Value = False Then
Cr.SelectionFormula = "{MastCust.Date}=Calendar3rdQtr"
End If
Cr.Action = 1
End Sub
Private Sub cmd4th_Click()
Call Report
If optMPS.Value = True Or optDPS.Value = True Or optTotPay.Value = True Then
Cr.SelectionFormula = "{PayList.Date}=Calendar4thQtr"
Cr.Action = 1
Exit Sub
End If

If optExpiry.Value = True Or optMExpiry.Value = True Then
Cr.SelectionFormula = "{MastCust.Expiry}=Calendar4thQtr"
ElseIf optExpiry.Value = False And optMExpiry.Value = False Then
Cr.SelectionFormula = "{MastCust.Date}=Calendar4thQtr"
End If
Cr.Action = 1
End Sub

Private Sub optTotPay_Click()
frmdet.Visible = True
frmview.Visible = True
txtRep.Text = "TotalPay"
End Sub

Private Sub TotalPay()
On Error GoTo errdps
Cr.Reset
Cr.ReportTitle = "Datewise Total Payment"
Cr.WindowState = crptMaximized
Cr.WindowTitle = "Datewise Total Payment"
Cr.WindowShowGroupTree = True
Cr.DataFiles(0) = App.Path & "\Data\Data97.mdb"
Cr.ReportFileName = App.Path & "\Reports\rptTotalPayments.rpt"

Cr.Destination = crptToWindow
Cr.WindowShowRefreshBtn = True
'Cr.Action = 1
Exit Sub
errdps:
MsgBox Err.Description, vbOKOnly, "MLM"
End Sub

Private Sub optTotPay_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub
