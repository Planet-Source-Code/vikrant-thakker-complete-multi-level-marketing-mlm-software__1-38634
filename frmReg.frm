VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmReg 
   Caption         =   "Member Registration .....     Real Money Zone"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   915
      Left            =   90
      TabIndex        =   53
      Top             =   7335
      Width           =   9015
      Begin VB.Frame Frmbutton 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   510
         Left            =   1395
         TabIndex        =   54
         Top             =   225
         Width           =   6495
         Begin VB.CommandButton cmdCancel 
            BackColor       =   &H0080C0FF&
            Caption         =   "&Cancel"
            Height          =   420
            Left            =   2790
            Style           =   1  'Graphical
            TabIndex        =   61
            Top             =   45
            Width           =   870
         End
         Begin VB.CommandButton cmdModify 
            BackColor       =   &H0080C0FF&
            Caption         =   "&Modify"
            Height          =   420
            Left            =   960
            Style           =   1  'Graphical
            TabIndex        =   60
            Top             =   45
            Width           =   915
         End
         Begin VB.CommandButton cmdSave 
            BackColor       =   &H0080C0FF&
            Caption         =   "&Save"
            Height          =   420
            Left            =   1890
            Style           =   1  'Graphical
            TabIndex        =   59
            Top             =   45
            Width           =   870
         End
         Begin VB.CommandButton cmdNext 
            BackColor       =   &H0080C0FF&
            Caption         =   "&Next"
            Height          =   420
            Left            =   5535
            Style           =   1  'Graphical
            TabIndex        =   58
            Top             =   45
            Width           =   870
         End
         Begin VB.CommandButton cmdPrev 
            BackColor       =   &H0080C0FF&
            Caption         =   "&Previous"
            Height          =   420
            Left            =   4590
            Style           =   1  'Graphical
            TabIndex        =   57
            Top             =   45
            Width           =   915
         End
         Begin VB.CommandButton cmdRemove 
            BackColor       =   &H0080C0FF&
            Caption         =   "&Remove"
            Height          =   420
            Left            =   3690
            Style           =   1  'Graphical
            TabIndex        =   56
            ToolTipText     =   "Remove facility is Not Available"
            Top             =   45
            Width           =   870
         End
         Begin VB.CommandButton cmdAdd 
            BackColor       =   &H0080C0FF&
            Caption         =   "&Add"
            Height          =   420
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   55
            Top             =   60
            Width           =   870
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   915
      Left            =   9135
      TabIndex        =   50
      Top             =   7335
      Width           =   2670
      Begin MLM.cmd cmdHelp 
         Height          =   465
         Left            =   180
         TabIndex        =   51
         Top             =   270
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
      Begin MLM.cmd cmdClose 
         Height          =   465
         Left            =   1395
         TabIndex        =   52
         Top             =   270
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
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H0041E9D8&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   5790
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdClose1 
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
      Left            =   10395
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   5985
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   10860
      TabIndex        =   46
      Top             =   60
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Frame frm2 
      Caption         =   "   Referrence"
      Height          =   1515
      Left            =   90
      TabIndex        =   35
      Top             =   5820
      Width           =   11745
      Begin VB.CommandButton cmdList 
         BackColor       =   &H0041E9D8&
         Caption         =   "&List"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   315
         Width           =   855
      End
      Begin VB.TextBox txtRefEmail 
         Appearance      =   0  'Flat
         BackColor       =   &H00C6D6B8&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   7620
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   18
         ToolTipText     =   $"frmReg.frx":0000
         Top             =   1020
         Width           =   3855
      End
      Begin VB.TextBox txtIntro 
         Appearance      =   0  'Flat
         BackColor       =   &H00C6D6B8&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1980
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   17
         ToolTipText     =   $"frmReg.frx":00C9
         Top             =   1020
         Width           =   3855
      End
      Begin VB.TextBox txtRefId 
         Appearance      =   0  'Flat
         BackColor       =   &H00C6D6B8&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1980
         MaxLength       =   15
         TabIndex        =   16
         ToolTipText     =   $"frmReg.frx":0193
         Top             =   360
         Width           =   2235
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Joining Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6180
         TabIndex        =   38
         Top             =   1020
         Width           =   1335
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Introduced by"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   300
         TabIndex        =   37
         Top             =   1080
         Width           =   3315
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Ref. I.D. No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   300
         TabIndex        =   36
         Top             =   420
         Width           =   1335
      End
   End
   Begin VB.Frame frm1 
      Caption         =   "  Member Details"
      Height          =   5775
      Left            =   90
      TabIndex        =   19
      Top             =   30
      Width           =   11745
      Begin VB.TextBox txtAmtPaid 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6AC7D&
         Enabled         =   0   'False
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
         Left            =   1980
         MaxLength       =   5
         TabIndex        =   15
         ToolTipText     =   $"frmReg.frx":024D
         Top             =   5070
         Width           =   1095
      End
      Begin VB.TextBox txtID 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6AC7D&
         Enabled         =   0   'False
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
         Left            =   1980
         MaxLength       =   15
         TabIndex        =   0
         ToolTipText     =   "IDNO of the member is created automatically by the computer. And you cannot change it. Each member has a unique IDNO."
         Top             =   480
         Width           =   3855
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6AC7D&
         Enabled         =   0   'False
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
         Left            =   1980
         MaxLength       =   25
         TabIndex        =   2
         ToolTipText     =   "Name of the Registered Member"
         Top             =   1140
         Width           =   3855
      End
      Begin VB.TextBox txtNom 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6AC7D&
         Enabled         =   0   'False
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
         Left            =   1980
         MaxLength       =   25
         TabIndex        =   4
         ToolTipText     =   "Enter the name of the nominee, to whom future dealings can be proceeded, incase of Member's unavailability"
         Top             =   1800
         Width           =   3855
      End
      Begin VB.TextBox txtCity 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6AC7D&
         Enabled         =   0   'False
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
         Left            =   6780
         MaxLength       =   15
         TabIndex        =   6
         ToolTipText     =   "City of the Member"
         Top             =   2430
         Width           =   1815
      End
      Begin VB.TextBox txtState 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6AC7D&
         Enabled         =   0   'False
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
         Left            =   9720
         MaxLength       =   15
         TabIndex        =   7
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox txtCountry 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6AC7D&
         Enabled         =   0   'False
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
         Left            =   6780
         MaxLength       =   15
         TabIndex        =   8
         Top             =   3060
         Width           =   1815
      End
      Begin VB.TextBox txtZip 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6AC7D&
         Enabled         =   0   'False
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
         Left            =   9720
         MaxLength       =   9
         TabIndex        =   9
         Top             =   3090
         Width           =   1695
      End
      Begin VB.TextBox txtTelR 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6AC7D&
         Enabled         =   0   'False
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
         Left            =   1980
         MaxLength       =   12
         TabIndex        =   10
         ToolTipText     =   "Member's Residence Telephone No."
         Top             =   3750
         Width           =   3855
      End
      Begin VB.TextBox txtTelO 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6AC7D&
         Enabled         =   0   'False
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
         Left            =   6750
         MaxLength       =   12
         TabIndex        =   11
         ToolTipText     =   "Member's Office Telephone No."
         Top             =   3720
         Width           =   1815
      End
      Begin VB.TextBox txtCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6AC7D&
         Enabled         =   0   'False
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
         Left            =   9720
         MaxLength       =   15
         TabIndex        =   12
         Top             =   3720
         Width           =   1695
      End
      Begin VB.TextBox txtDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6AC7D&
         Enabled         =   0   'False
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
         Left            =   9060
         TabIndex        =   1
         Top             =   540
         Width           =   1515
      End
      Begin VB.TextBox txtDOB 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6AC7D&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd MMM yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Enabled         =   0   'False
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
         Left            =   9060
         TabIndex        =   3
         Top             =   1140
         Width           =   2355
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6AC7D&
         Enabled         =   0   'False
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
         Left            =   7680
         MaxLength       =   50
         TabIndex        =   13
         Top             =   5070
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.TextBox txtMarr 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6AC7D&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd MMM yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Enabled         =   0   'False
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
         Left            =   1980
         MaxLength       =   8
         TabIndex        =   14
         ToolTipText     =   "Marriage date of Member if married..."
         Top             =   4380
         Width           =   3855
      End
      Begin RichTextLib.RichTextBox txtAdd 
         Height          =   915
         Left            =   2010
         TabIndex        =   5
         ToolTipText     =   "Address of the Registered Member"
         Top             =   2490
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   1614
         _Version        =   393217
         BackColor       =   15117437
         Enabled         =   0   'False
         MaxLength       =   200
         Appearance      =   0
         TextRTF         =   $"frmReg.frx":02ED
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label23 
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   5100
         Width           =   1335
      End
      Begin VB.Label Label22 
         Caption         =   "DD/MM/YY"
         Height          =   315
         Left            =   10680
         TabIndex        =   42
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   8880
         TabIndex        =   41
         Top             =   540
         Width           =   315
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   1800
         TabIndex        =   40
         Top             =   1200
         Width           =   315
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   1800
         TabIndex        =   39
         Top             =   600
         Width           =   315
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "ID No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   540
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Name Mr./Ms."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Nominee's Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   240
         TabIndex        =   32
         Top             =   1860
         Width           =   1875
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Address "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6060
         TabIndex        =   30
         Top             =   2460
         Width           =   1395
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   9000
         TabIndex        =   29
         Top             =   2460
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6000
         TabIndex        =   28
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Zip"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   9000
         TabIndex        =   27
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Tel. (R)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   3780
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "(O)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6000
         TabIndex        =   25
         Top             =   3780
         Width           =   1335
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Cell"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   9000
         TabIndex        =   24
         Top             =   3780
         Width           =   1335
      End
      Begin VB.Label Label12 
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7680
         TabIndex        =   23
         Top             =   540
         Width           =   1335
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Birth"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7680
         TabIndex        =   22
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Email Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6030
         TabIndex        =   21
         Top             =   5100
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Marriage Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   4410
         Width           =   1335
      End
   End
   Begin MSDataGridLib.DataGrid dg 
      Height          =   5625
      Left            =   2520
      TabIndex        =   45
      Top             =   30
      Visible         =   0   'False
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   9922
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   15499943
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
   Begin VB.Label lblID 
      BackColor       =   &H00C0C0FF&
      Height          =   315
      Left            =   10530
      TabIndex        =   44
      Top             =   6975
      Visible         =   0   'False
      Width           =   1515
   End
End
Attribute VB_Name = "frmReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim rsMemList As Recordset
Dim Main, TotalAmt, Level0, Level1, Level2, Level3, Level4, Total As Integer
Dim cnt, Level, comm, diff As Integer
Dim dd, mm, yy, d, m, y9, y10
Dim Modify As Boolean

Private Sub cmdAdd_Click()
On Error GoTo aerr
Modify = False
Call Modi
'txtAmtPaid.Enabled = True
txtID.Enabled = True
txtDate.Enabled = True
txtName.Enabled = True
txtDOB.Enabled = True
txtNom.Enabled = True
txtAdd.Enabled = True
txtCity.Enabled = True
txtState.Enabled = True
txtCountry.Enabled = True
txtZip.Enabled = True
txtTelR.Enabled = True
txtTelO.Enabled = True
txtCell.Enabled = True
txtEmail.Enabled = True
txtMarr.Enabled = True
txtRefId.Enabled = True
txtIntro.Enabled = True
txtRefEmail.Enabled = True
cmdList.Enabled = True
cmdOK.Enabled = True
frm2.Enabled = True

'txtID.Text = ""

'If rsMem.RecordCount = 0 Then
'txtID.Text = "1"
'Else
'rsMem.MoveLast
'txtID.Text = rsMem!IdNo + 1
'End If


txtID.Text = rsMem.RecordCount
check:
If rsMem.RecordCount > 0 Then
rsMem.MoveFirst
    For i = 0 To rsMem.RecordCount - 1 Step 1
        If rsMem!IDNO = txtID.Text Then
           ' MsgBox "ID already Exits !"
            txtID.Text = (txtID.Text) + 1
            GoTo check
        End If
        rsMem.MoveNext
    Next
End If
            

txtID.Locked = True
txtDate.Text = ""
txtName.Text = ""
txtDOB.Text = ""
txtNom.Text = ""
txtAdd.Text = ""
txtCity.Text = ""
txtState.Text = ""
txtCountry.Text = ""
txtZip.Text = ""
txtTelR.Text = ""
txtTelO.Text = ""
txtCell.Text = ""
txtEmail.Text = ""
txtMarr.Text = ""
txtRefId.Text = ""
txtIntro.Text = ""
txtRefEmail.Text = ""
txtState.Text = rsDef!defaultstate
txtCity.Text = rsDef!defaultCity
txtCountry.Text = rsDef!defaultcountry
txtAmtPaid.Text = rsDef!defaultamt
If rsDef!defaultdate = "Y" Then
    txtDate.Text = Date
Else
txtDate.Text = ""
End If
rsMem.AddNew

cmdCancel.Enabled = True
cmdSave.Enabled = True
cmdModify.Enabled = False
cmdRemove.Enabled = False
cmdPrev.Enabled = False
cmdNext.Enabled = False
cmdAdd.Enabled = False
cmdClose.Enabled = False

'txtID.SetFocus
txtDate.SetFocus
cmdSave.Enabled = False
Exit Sub
aerr:
MsgBox Err.Description, vbOKOnly, "MLM"
End Sub

Private Sub cmdCancel_Click()
On Error GoTo cerr
Modify = False
Call Modi
rsMem.CancelUpdate
txtID.Enabled = False
txtDate.Enabled = False
txtName.Enabled = False
txtDOB.Enabled = False
txtNom.Enabled = False
txtAdd.Enabled = False
txtCity.Enabled = False
txtState.Enabled = False
txtCountry.Enabled = False
txtZip.Enabled = False
txtTelR.Enabled = False
txtTelO.Enabled = False
txtCell.Enabled = False
txtEmail.Enabled = False
txtMarr.Enabled = False
txtRefId.Enabled = False
txtIntro.Enabled = False
txtRefEmail.Enabled = False
frm2.Enabled = True
cmdAdd.Enabled = True
cmdModify.Enabled = True
cmdPrev.Enabled = True
cmdNext.Enabled = True
'cmdRemove.Enabled = True   origninal
cmdRemove.Enabled = False
cmdCancel.Enabled = False
cmdSave.Enabled = False
cmdClose.Enabled = True

Call cmdPrev_Click
Exit Sub
cerr:
MsgBox Err.Description, vbOKOnly, "MLM"
End Sub

Private Sub cmdClose_Click()
On Error GoTo eerr
frmMain.Show
Unload Me

'    Ans = MsgBox("Are you sure You want to quit ?", vbYesNo, "Warning")
'    If Ans = vbYes Then
'    End
'    Else
'    Exit Sub
'    End If
Exit Sub
eerr:
MsgBox Err.Description, vbOKOnly, "MLM"
End Sub

Private Sub cmdHelp_Click()
MsgBox "Details of the new registered member are supposed to be entered over here. You can also edit/delete the information of previous created members. IDNO, Name, and Date are compulsory fields, and these fields entered once cannot be changed latter. So special care needs to be taken before entering this information. Certain fields like City, Country, etc. may show some default text. This text comes from the information entered in the Default form. You can change this manually if required. All fields are self explanatory, and donot need any further explanations.", vbOKOnly, "Help"
End Sub

Private Sub cmdList_Click()
On Error GoTo lerr
rsMemList.Requery
If rsMemList.RecordCount = 0 Then
MsgBox "No Records Found !", vbOKOnly, "MLM"
If txtRefId.Enabled = True Then txtRefId.SetFocus
Exit Sub
End If

If rsMemList.RecordCount > 0 Then
dg.Refresh
dg.Visible = True
cmdList.Visible = False
Frmbutton.Visible = False
cmdClose.Visible = False
cmdOK.Visible = True

frm1.Visible = False
frm2.Visible = False

End If


dg.Row = 0
'MsgBox "approxCount : " & dg.ApproxCount
'MsgBox "Reccount : " & rsMem.RecordCount
For i = 0 To dg.ApproxCount - 1 Step 1
   ' MsgBox (dg.Columns(0).Text)
    If dg.Columns(0).Text = txtRefId.Text Then
       ' dg.Row = I
        'GoTo jump
        Exit Sub
        'GoTo jump
    End If
   
   If i < dg.ApproxCount Then dg.Row = dg.Row + 1
'    dg.Row = I
    'dg.Row = I + 1
    If i = dg.ApproxCount Then
    MsgBox "Error", vbOKOnly, "MLM"
    Exit Sub
    End If
    
Next
jump:

Exit Sub
lerr:
'MsgBox Err.Description, vbOKOnly, "MLM"
End Sub

Private Sub cmdModify_Click()
On Error GoTo merr
Modify = True
Call Modi
'txtID.Enabled = True
'txtDate.Enabled = True
'txtName.Enabled = True
txtDOB.Enabled = True
txtNom.Enabled = True
txtAdd.Enabled = True
txtCity.Enabled = True
txtState.Enabled = True
txtCountry.Enabled = True
txtZip.Enabled = True
txtTelR.Enabled = True
txtTelO.Enabled = True
txtCell.Enabled = True
txtEmail.Enabled = True
txtMarr.Enabled = True
txtRefId.Enabled = True
txtIntro.Enabled = True
txtRefEmail.Enabled = True
frm2.Enabled = False

cmdCancel.Enabled = True
cmdSave.Enabled = True
cmdPrev.Enabled = False
cmdNext.Enabled = False
cmdAdd.Enabled = False
cmdRemove.Enabled = False
cmdModify.Enabled = False
cmdClose.Enabled = False

'txtID.SetFocus
Exit Sub
merr:
MsgBox Err.Description, vbOKOnly, "MLM"
End Sub

Private Sub cmdNext_Click()
On Error GoTo nerr
Modify = False
Call Modi
frm2.Enabled = True
If rsMem.RecordCount = 0 Then
    cmdModify.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdRemove.Enabled = False
    cmdCancel.Enabled = False
    Exit Sub
End If
rsMem.MoveNext
If rsMem.EOF = True Then rsMem.MoveLast
    showall
Exit Sub
nerr:
MsgBox Err.Description, vbOKOnly, "MLM"
End Sub

Private Sub cmdOK_Click()
On Error GoTo oerr
frm2.Enabled = True
dg.Visible = False
Frmbutton.Visible = True
cmdClose.Visible = True
cmdOK.Visible = False

frm1.Visible = True
cmdList.Visible = True
frm2.Visible = True

Exit Sub
oerr:
MsgBox Err.Description, vbOKOnly, "MLM"
End Sub

Private Sub cmdPrev_Click()
On Error GoTo perr
Modify = False
Call Modi
frm2.Enabled = True
If rsMem.RecordCount = 0 Then
    cmdModify.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdRemove.Enabled = False
    cmdCancel.Enabled = False
    Exit Sub
End If
If rsMem.BOF = False And rsMem.EOF = False Then rsMem.MovePrevious
If rsMem.BOF = True Then rsMem.MoveFirst
showall
   ' cmdRemove.Enabled = True   original
    cmdRemove.Enabled = False

Exit Sub
perr:
MsgBox Err.Description, vbOKOnly, "MLM"
End Sub

Private Sub cmdRemove_Click()
On Error GoTo rerr
MsgBox "Remove Facility Not Available !", vbOKOnly, "MLM"
'Modify = False
'Call Modi
'frm2.Enabled = True
' If rsMem.RecordCount = 0 Then
'    cmdModify.Enabled = False
'    cmdPrev.Enabled = False
'    cmdNext.Enabled = False
'    cmdRemove.Enabled = False
'    cmdCancel.Enabled = False
'    Exit Sub
'End If

'       rsMem.Delete
'       If rsMem.EOF = False Then rsMem.MoveNext
'       showall
Exit Sub
rerr:
MsgBox Err.Description, vbOKOnly, "MLM"
End Sub

Private Sub cmdSave_Click()
'On Error GoTo serr

If Trim(txtID.Text) = "" Or Trim(txtName.Text) = "" Or Trim(txtDate.Text) = "" Then
    Exit Sub
End If
frm2.Enabled = True

'Call TotRef
If rsMemList.RecordCount > 0 Then rsMemList.MoveFirst
'
         Do While rsMemList.EOF = False
    If txtRefId.Text = rsMemList!IDNO Or Trim(txtRefId.Text) = "" Then
'        rsMemList!TotRef = rsMemList!TotRef + 1
'        rsMemList.Update
'        'MsgBox "RefId Found !"
         Exit Do
    End If
            rsMemList.MoveNext
        If rsMemList.EOF = True Then
            MsgBox "Select the RefId from Help List !", vbOKOnly, "MLM"
            txtRefId.SetFocus
            Exit Sub
        End If
    Loop
    'End If




If txtID.Text <> "" Then rsMem!IDNO = txtID.Text
If txtName.Text <> "" Then rsMem!Name = txtName.Text
If txtDate.Text <> "" Then rsMem!Date = txtDate.Text
If txtDOB.Text <> "" Then rsMem!dob = txtDOB.Text
If txtNom.Text <> "" Then rsMem!nom = txtNom.Text
If txtAdd.Text <> "" Then rsMem!address = txtAdd.Text
If txtCity.Text <> "" Then rsMem!city = txtCity.Text
If txtState.Text <> "" Then rsMem!State = txtState.Text
If txtCountry.Text <> "" Then rsMem!country = txtCountry.Text
If txtZip.Text <> "" Then rsMem!zip = txtZip.Text
If txtTelR.Text <> "" Then rsMem!telR = txtTelR.Text
If txtTelO.Text <> "" Then rsMem!telo = txtTelO.Text
If txtCell.Text <> "" Then rsMem!cell = txtCell.Text
If txtEmail.Text <> "" Then rsMem!email = txtEmail.Text
If txtMarr.Text <> "" Then rsMem!marr = txtMarr.Text
If txtRefId.Text <> "" Then rsMem!RefId = txtRefId.Text
If txtIntro.Text <> "" Then rsMem!Intro = txtIntro.Text
If txtRefEmail.Text <> "" Then rsMem!RefEmail = txtRefEmail.Text
If txtAmtPaid.Text <> "" Then rsMem!AmtPaid = txtAmtPaid.Text

rsMem!expiry = DateAdd("M", rsDef!DefaultExpiry, txtDate.Text)

If Year(txtDate.Text) < Year(rsExpiry!CurrentYear) Then
    MsgBox "You cannot work in previous working year ! Contact Developer", vbCritical, "MLM"
Exit Sub
ElseIf Year(txtDate.Text) > Year(rsExpiry!CurrentYear) Then
    MsgBox "You cannot work in next working year ! Contact Developer", vbCritical, "MLM"
Exit Sub
End If

If rsMem.RecordCount > 0 Then
    rsMem.MoveFirst
    Do While rsMem.EOF = False
    If txtRefId.Text = rsMem!IDNO Or Trim(txtRefId.Text) = "" Then
    ' MsgBox "RefId Found !"
                Exit Do
    End If
    rsMem.MoveNext
    If rsMem.EOF = True Then
   ' MsgBox "Please Check The RefId !"
    Exit Sub
    End If
    Loop
    End If
    
'Call expiryDate

rsMem.Update
rsMemList.Requery

If Modify = False Then
    Call LevelwiseEntry
    Call TotRef
End If

cmdSave.Enabled = False
cmdAdd.Enabled = True
cmdModify.Enabled = True
cmdPrev.Enabled = True
cmdNext.Enabled = True
cmdCancel.Enabled = False
'cmdRemove.Enabled = True  original
cmdRemove.Enabled = False
cmdClose.Enabled = True


txtID.Enabled = True
txtDate.Enabled = True
txtName.Enabled = True
txtDOB.Enabled = True
txtNom.Enabled = True
txtAdd.Enabled = True
txtCity.Enabled = True
txtState.Enabled = True
txtCountry.Enabled = True
txtZip.Enabled = True
txtTelR.Enabled = True
txtTelO.Enabled = True
txtCell.Enabled = True
txtEmail.Enabled = True
txtMarr.Enabled = True
txtRefId.Enabled = True
txtIntro.Enabled = True
txtRefEmail.Enabled = True

txtID.Text = ""
txtDate.Text = ""
txtName.Text = ""
txtDOB.Text = ""
txtNom.Text = ""
txtAdd.Text = ""
txtCity.Text = ""
txtState.Text = ""
txtCountry.Text = ""
txtZip.Text = ""
txtTelR.Text = ""
txtTelO.Text = ""
txtCell.Text = ""
txtEmail.Text = ""
txtMarr.Text = ""
txtRefId.Text = ""
txtIntro.Text = ""
txtRefEmail.Text = ""
'txtID.SetFocus

If cmdAdd.Enabled = True Then cmdAdd.SetFocus
Exit Sub
serr:
MsgBox Err.Description, vbOKOnly, "MLM"

txtID.Enabled = True
txtDate.Enabled = True
txtName.Enabled = True
txtDOB.Enabled = True
txtNom.Enabled = True
txtAdd.Enabled = True
txtCity.Enabled = True
txtState.Enabled = True
txtCountry.Enabled = True
txtZip.Enabled = True
txtTelR.Enabled = True
txtTelO.Enabled = True
txtCell.Enabled = True
txtEmail.Enabled = True
txtMarr.Enabled = True
txtRefId.Enabled = True
txtIntro.Enabled = True
txtRefEmail.Enabled = True

txtID.Text = ""
txtDate.Text = ""
txtName.Text = ""
txtDOB.Text = ""
txtNom.Text = ""
txtAdd.Text = ""
txtCity.Text = ""
txtState.Text = ""
txtCountry.Text = ""
txtZip.Text = ""
txtTelR.Text = ""
txtTelO.Text = ""
txtCell.Text = ""
txtEmail.Text = ""
txtMarr.Text = ""
txtRefId.Text = ""
txtIntro.Text = ""
txtRefEmail.Text = ""
txtID.SetFocus
End Sub

Private Sub TotRef()
' To update the total no of references given by any Member
If rsMem.RecordCount > 0 Then rsMem.MoveFirst

         Do While rsMem.EOF = False
    If txtRefId.Text = rsMem!IDNO Or Trim(txtRefId.Text) = "" Then
        rsMem!TotRef = rsMem!TotRef + 1
        rsMem.Update
        'MsgBox "RefId Found !"
         Exit Do
    End If
            rsMem.MoveNext
        If rsMem.EOF = True Then
            MsgBox "Select the RefId from Help List !", vbOKOnly, "MLM"
            txtRefId.SetFocus
            Exit Sub
        End If
    Loop
    'End If
End Sub


Private Sub dg_DblClick()
On Error GoTo dgcerr
txtRefId.Text = (dg.Columns(0).Text)
txtIntro.Text = dg.Columns(1).Text
txtRefEmail.Text = dg.Columns(2).Text
Call cmdOK_Click
Exit Sub
dgcerr:
MsgBox Err.Description, vbOKOnly, "MLM"
End Sub

Private Sub dg_KeyPress(KeyAscii As Integer)
On Error GoTo dgkerr
If KeyAscii = 13 Then
txtRefId.Text = (dg.Columns(0).Text)
txtIntro.Text = dg.Columns(1).Text
txtRefEmail.Text = dg.Columns(2).Text


frm2.Enabled = True
dg.Visible = False
Frmbutton.Visible = True
cmdClose.Visible = True
cmdOK.Visible = False

frm1.Visible = True
cmdList.Visible = True
frm2.Visible = True
'Call cmdOK_Click
End If
Exit Sub
dgkerr:
MsgBox Err.Description, vbOKOnly, "MLM"
End Sub

Private Sub dg_LostFocus()
On Error GoTo dglerr

If dg.Visible = False And txtRefEmail.Enabled = True Then
    txtRefEmail.SetFocus
End If
Exit Sub
dglerr:
MsgBox Err.Description, vbOKOnly, "MLM"
End Sub

Private Sub Form_Activate()
cmdAdd.SetFocus
cmdList.ToolTipText = "To view the List of the Registered Members with less then 4 members under their name. As no member can refer to more then 3 members. You can select the referrence IDNo, Name and Date with the help of this list. To select the IDNO, dblclick or press enter key on the selected IDNo, and all the information will be entered in these fields. Note again : maximum 3 members can be created by a member. i.e. John can refer to only 3 members A,B and C. And as soon as 3 members are registered by any member, his name will not be visible under this list."
End Sub

Private Sub Form_Load()
On Error GoTo ferr
'If rsDef!defaultdate = "Y" Then
'    txtDate.Text = Date
'Else
'    txtDate.Text = ""
'End If

txtState.Text = rsDef!defaultstate
txtCity.Text = rsDef!defaultCity
txtCountry.Text = rsDef!defaultcountry
txtAmtPaid.Text = rsDef!defaultamt
Main = rsDef!Main
TotalAmt = rsDef!TotalAmt
If rsDef!defaultdate = "Y" Then
    txtDate.Text = Date
Else
txtDate.Text = ""
End If



' If rsMem.RecordCount = 0 Then
    cmdModify.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdRemove.Enabled = False
    cmdCancel.Enabled = False
    
 ' End If
    
If rsMem.RecordCount > 0 Then
    cmdAdd.Enabled = True
    cmdPrev.Enabled = True
    cmdNext.Enabled = True
End If
    cmdSave.Enabled = False
    
If rsComm.RecordCount > 0 Then rsComm.MoveFirst

Level0 = rsComm!comm    'Level = 50
If rsComm.EOF <> True Then rsComm.MoveNext
Level1 = rsComm!comm    'Level1 = 100
If rsComm.EOF <> True Then rsComm.MoveNext
Level2 = rsComm!comm    'Level2 = 50
If rsComm.EOF <> True Then rsComm.MoveNext
Level3 = rsComm!comm    'Level3 = 25
If rsComm.EOF <> True Then rsComm.MoveNext
Level4 = rsComm!comm    'Level4 = 25

' This is how Rs. 300 are distributed

Set rsMemList = New ADODB.Recordset
rsMemList.Open "select IdNo,Name,Date,TotRef from MastCust where MastCust!TotRef<3", conn, adOpenStatic, adLockOptimistic

Set dg.DataSource = rsMemList

'dg.ScrollBars = dbgAutomatic
Exit Sub
ferr:
MsgBox Err.Description, vbOKOnly, "MLM"
End Sub

Private Sub txtAdd_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtAdd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtAdd_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub txtAmtPaid_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtAmtPaid_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtAmtPaid_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub txtCell_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtCell_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtCell_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub txtCity_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtCity_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtCity_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub txtCountry_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtCountry_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtCountry_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
On Error GoTo derr
Select Case KeyAscii
    
    Case vbKeyBack, 48 To 57, vbKeyReturn, 47, vbKeyEscape
    Case Else
         MsgBox "No Special Characters are allowed! Please enter Numbers and / only! ", vbOKOnly, "MLM"
         KeyAscii = 0
         txtDate.SetFocus
        
         Exit Sub
    End Select

If KeyAscii = 13 Then
   datevali (txtDate.Text)
End If

'If KeyAscii = 13 Then
'    SendKeys "{TAB}"
'    KeyAscii = 0
'End If
Exit Sub
derr:
MsgBox Err.Description, vbOKOnly, "MLM"
End Sub
Private Sub txtDOB_KeyPress(KeyAscii As Integer)
On Error GoTo doberr
Select Case KeyAscii
    
    Case vbKeyBack, 48 To 57, vbKeyReturn, 47, vbKeyEscape
    Case Else
         MsgBox "No Special Characters are allowed! Please enter Numbers and / only! "
         KeyAscii = 0
         txtDOB.SetFocus
        
         Exit Sub
    End Select
    
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
Exit Sub
doberr:
MsgBox Err.Description, vbOKOnly, "MLM"
End Sub

Private Sub txtEmail_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtEmail_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtGuj_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtEmail_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub txtID_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtID_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub txtIntro_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtIntro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtIntro_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub txtMarr_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtMarr_KeyPress(KeyAscii As Integer)
On Error GoTo merr
Select Case KeyAscii
    
    Case vbKeyBack, 48 To 57, vbKeyReturn, 47, vbKeyEscape
    Case Else
         MsgBox "No Special Characters are allowed! Please enter Numbers and / only! "
         KeyAscii = 0
         txtMarr.SetFocus
        
         Exit Sub
    End Select
    
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
Exit Sub
merr:
MsgBox Err.Description, vbOKOnly, "MLM"

End Sub

Private Sub txtMarr_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub txtName_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtName_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub txtNom_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtNom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub
Private Sub txtO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtNom_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub


Private Sub txtRefEmail_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtRefEmail_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If cmdSave.Enabled = True Then
    cmdSave.SetFocus
    Exit Sub
    End If
    SendKeys "{TAB}"
    KeyAscii = 0
End If
'If cmdSave.Enabled = True Then cmdSave.SetFocus
End Sub

Private Sub txtRefEmail_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub txtRefId_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtRefId_KeyPress(KeyAscii As Integer)
On Error GoTo refkerr
If KeyAscii = 13 Then
    If Trim(txtRefId.Text) <> "" Then
        Call RefIdcheck
        Call cmdList_Click
    Else
    SendKeys "{TAB}"
    KeyAscii = 0
    End If
End If

Exit Sub
refkerr:
MsgBox Err.Description, vbOKOnly, "MLM"
End Sub
Private Sub RefIdcheck()
On Error GoTo refiderr
If rsRefIdCheck.RecordCount > 0 Then
    rsRefIdCheck.MoveFirst
    For i = 0 To rsRefIdCheck.RecordCount - 1 Step 1
        If txtRefId.Text = rsRefIdCheck!IDNO Then
            Call cmdList_Click
        '    GoTo jum
        Exit Sub
        End If
        If rsRefIdCheck.EOF = True Then
            MsgBox "Please Check the Ref.Id !"
            txtRefId.Text = ""
            Exit Sub
        End If
    Next
End If
jum:
'n = rsMem.Bookmark
'    If rsRefIdCheck.RecordCount > 0 Then
'n = rsRefIdCheck.Bookmark
'        rsRefIdCheck.MoveFirst
'        Do While rsRefIdCheck.EOF = False
'        If txtRefId.Text = rsRefIdCheck!IdNo Then ' Or Trim(txtRefId.Text) = "" Then
'                Call cmdList_Click
'                Exit Do
'        End If
'        rsRefIdCheck.MoveNext
'        If rsRefIdCheck.EOF = True Then
'        MsgBox "Please Check The RefId !"
'        txtRefId.Text = ""
'        Exit Sub
'        End If
'        Loop
'    End If
'rsRefIdCheck.Bookmark = n
Exit Sub
refiderr:
MsgBox Err.Description, vbOKOnly, "MLM"
End Sub

Private Sub txtRefId_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub txtState_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtState_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtState_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub txtTelO_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtTelO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtTelO_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub txtTelR_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtTelR_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtTelR_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub txtZip_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtZip_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub showall()
On Error GoTo showerr
txtID.Text = ""
txtDate.Text = ""
txtName.Text = ""
txtDOB.Text = ""
txtNom.Text = ""
txtAdd.Text = ""
txtCity.Text = ""
txtState.Text = ""
txtCountry.Text = ""
txtZip.Text = ""
txtTelR.Text = ""
txtTelO.Text = ""
txtCell.Text = ""
txtEmail.Text = ""
txtMarr.Text = ""
txtRefId.Text = ""
txtIntro.Text = ""
txtRefEmail.Text = ""

If rsMem.RecordCount = 0 Then
    cmdModify.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdRemove.Enabled = False
    Exit Sub
End If
If Not IsNull(rsMem!Name) Then txtName.Text = rsMem!Name
If Not IsNull(rsMem!IDNO) Then txtID.Text = rsMem!IDNO
If Not IsNull(rsMem!Date) Then txtDate.Text = rsMem!Date
If Not IsNull(rsMem!dob) Then txtDOB.Text = rsMem!dob
If Not IsNull(rsMem!nom) Then txtNom.Text = rsMem!nom
If Not IsNull(rsMem!address) Then txtAdd.Text = rsMem!address
If Not IsNull(rsMem!city) Then txtCity.Text = rsMem!city
If Not IsNull(rsMem!State) Then txtState.Text = rsMem!State
If Not IsNull(rsMem!country) Then txtCountry.Text = rsMem!country
If Not IsNull(rsMem!zip) Then txtZip.Text = rsMem!zip
If Not IsNull(rsMem!telR) Then txtTelR.Text = rsMem!telR
If Not IsNull(rsMem!telo) Then txtTelO.Text = rsMem!telo
If Not IsNull(rsMem!cell) Then txtCell.Text = rsMem!cell
If Not IsNull(rsMem!email) Then txtEmail.Text = rsMem!email
If Not IsNull(rsMem!marr) Then txtMarr.Text = rsMem!marr
If Not IsNull(rsMem!RefId) Then txtRefId.Text = rsMem!RefId
If Not IsNull(rsMem!RefEmail) Then txtRefEmail.Text = rsMem!RefEmail
If Not IsNull(rsMem!Intro) Then txtIntro.Text = rsMem!Intro
If Not IsNull(rsMem!AmtPaid) Then txtAmtPaid.Text = rsMem!AmtPaid

Exit Sub
showerr:
MsgBox Err.Description, vbOKOnly, "MLM"
End Sub


Private Sub txt_GotFocus()
On Error GoTo focerr
'If addmod = 1 Then
    If Trim(txtName.Text) = "" Or Trim(txtID.Text) = "" Or Trim(txtDate.Text) = "" Then
        cmdSave.Enabled = False
        txtRefId.Enabled = False
        txtRefId.Text = ""
        txtRefEmail.Text = ""
        txtIntro.Text = ""
    Else
        cmdSave.Enabled = True
        txtRefId.Enabled = True
    End If
'End If
Exit Sub
focerr:
MsgBox Err.Description, vbOKOnly, "MLM"
End Sub

Private Sub txtZip_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Function datevali(dtt)
On Error GoTo dvalerr
d1 = 0
dd = 0
m1 = 0
mm = 0
Y1 = 0
yy = 0

d1 = InStr(1, dtt, "/")
If d1 > 0 Then
    dd = Mid(dtt, 1, d1 - 1)
Else
    MsgBox "Please enter / after date"
    txtDate.SetFocus
End If
    dlen = Len(dtt)
    dlen = dlen - d1
    mmid = Mid(dtt, d1 + 1, dlen)
    m1 = InStr(1, mmid, "/")
If m1 > 0 Then
    mm = Mid(mmid, 1, m1 - 1)
Else
    MsgBox "Please enter / after month"
    txtDate.SetFocus
End If
    dlen = Len(mmid)
    dlen = dlen - m1
    yy = Mid(mmid, m1 + 1, dlen)
If Len(dd) > 2 Then
    MsgBox "Please enter date of two digit"
    txtDate.SetFocus
ElseIf Val(dd) > 31 Or Val(dd) < 1 Then
    MsgBox "Plz enter date of less then 31"
    txtDate.SetFocus
ElseIf Val(mm) > 12 Or Val(mm) < 1 Then
    MsgBox "Plz enter month of less/equal then 12"
    txtDate.SetFocus
ElseIf Len(mm) > 2 Then
    MsgBox "Plz enter month of two digit"
    txtDate.SetFocus
ElseIf Len(yy) <> 2 And Len(yy) <> 4 Then
    MsgBox "Plz enter year of 2 or 4 digit"
    txtDate.SetFocus
ElseIf (Val(yy) < 1) Then
    MsgBox "Please enteryear Between financial year"
    txtDate.SetFocus
Else
    SendKeys "{TAB}"
    KeyAscii = 0
End If

Exit Function
dvalerr:
MsgBox Err.Description, vbOKOnly, "MLM"
End Function

Private Sub LevelwiseEntry()
'On Error GoTo lwerr
' cnt is a counter used to check that the commission is not added at a 0 level, but it is added from 1 to 4 Levels only
cnt = 0
Level = 0
TOT = 0
lblID.Caption = txtID.Text
again:
If txtRefId.Text = "" Then GoTo AdminAccount
If txtID.Text = "" Then GoTo AdminAccount

If Level = 1 Then comm = Level1
If Level = 2 Then comm = Level2
If Level = 3 Then comm = Level3
If Level = 4 Then comm = Level4
If Level < 1 Or Level > 4 Then comm = 0

If rsMem.RecordCount > 0 Then rsMem.MoveFirst

For i = 0 To rsMem.RecordCount - 1 Step 1
    
    If txtID.Text = rsMem!IDNO Then
        If Level > 0 And Level < 5 Then
                        If rsMem!expiry > Date Then
                Call AddAccount
                rsMem!Due = rsMem!Due + comm
                rsMem!Balance = rsMem!Balance + comm
                TOT = TOT + comm
            Else
                rsMem!expired = "Y"
                rsMem.Update
                If rsMem!RefId <> "" Then
                       txtID.Text = rsMem!RefId
                       GoTo again
                Else
                        txtID.Text = ""
                End If
            End If
        End If
        
        cnt = cnt + 1
        Level = Level + 1
    
        
        
           If rsMem!RefId <> "" Then
            txtID.Text = rsMem!RefId
               GoTo again
           Else
               txtID.Text = ""
            End If
  
    End If

rsMem.MoveNext
Next
'over:

AdminAccount:
diff = TotalAmt - TOT
If rsMem.RecordCount > 0 Then rsMem.MoveFirst
    For a = 0 To rsMem.RecordCount - 1 Step 1
        If rsMem!IDNO = Main Then
            rsMem!Balance = rsMem!Balance + diff
            rsMem!Due = rsMem!Due + diff
            rsMem.Update
        
        
        
        rsAcc!com = diff + comm
        rsAcc.Update
            Exit Sub
        End If
        rsMem.MoveNext
    Next
Exit Sub
lwerr:
MsgBox Err.Description, vbOKOnly, "MLM"
End Sub


Private Sub AddAccount()
On Error GoTo adderr
rsAcc.AddNew
rsAcc!Id = rsMem!IDNO
rsAcc!Name = rsMem!Name
rsAcc!Level = Level
rsAcc!RefId = lblID.Caption
rsAcc!RefName = txtName.Text
rsAcc!com = comm
rsAcc.Update
'n = rsAcc.Bookmark

' Here we want to add the exact amount that is added to the Admin Account by giving Referrence to this Member

'rsAcc!comm = diff
'rsAcc.Update
'GoTo SaveRecord
Exit Sub
adderr:
MsgBox Err.Description, vbOKOnly, "MLM"
End Sub

Private Sub Modi()
On Error GoTo merr
If Modify = True Then   ' If cmdmodify is clicked then
frm2.Enabled = False
txtID.Enabled = False
txtName.Enabled = False
txtDate.Enabled = False

ElseIf Modify = False Then  ' If cmdmodify is not clicked then
frm2.Enabled = True
txtID.Enabled = True
txtName.Enabled = True
txtDate.Enabled = True
End If
Exit Sub
merr:
MsgBox Err.Description, vbOKOnly, "MLM"
End Sub

