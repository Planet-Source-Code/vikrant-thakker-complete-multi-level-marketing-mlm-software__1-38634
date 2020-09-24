VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "find the top 4 Level of a particular User"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Final"
      Height          =   555
      Left            =   3060
      TabIndex        =   6
      Top             =   720
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   180
      TabIndex        =   1
      Top             =   2220
      Width           =   1635
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   555
      Left            =   2280
      TabIndex        =   0
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblID 
      BackColor       =   &H00C0C0FF&
      Height          =   315
      Left            =   180
      TabIndex        =   7
      Top             =   2700
      Width           =   1635
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   375
      Left            =   900
      TabIndex        =   5
      Top             =   60
      Width           =   1875
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   315
      Left            =   840
      TabIndex        =   4
      Top             =   600
      Width           =   1875
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   315
      Left            =   840
      TabIndex        =   3
      Top             =   1080
      Width           =   1755
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   315
      Left            =   840
      TabIndex        =   2
      Top             =   1560
      Width           =   1755
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Level1, Level2, Level3, Level4 As Integer
Dim cnt, Level, comm As Integer
Private Sub Command1_Click()

If rsMem.RecordCount > 0 Then rsMem.MoveFirst
no = 0
For i = 0 To rsMem.RecordCount - 1 Step 1
    If Text1.Text = rsMem!Idno Then
      Label1.Caption = rsMem!RefId
    no = no + 1
    GoTo Level2
    End If
    rsMem.MoveNext
Next
    
Level2:
rsMem.MoveFirst
For i = 0 To rsMem.RecordCount - 1 Step 1
    If Label1.Caption = rsMem!Idno Then
    Label2.Caption = rsMem!RefId
    no = no + 1
    GoTo Level3
    End If
    rsMem.MoveNext
Next

Level3:
rsMem.MoveFirst
For i = 0 To rsMem.RecordCount - 1 Step 1
    If Label2.Caption = rsMem!Idno Then
     Label3.Caption = rsMem!RefId
     no = no + 1
    GoTo Level4
    End If
    rsMem.MoveNext
Next

Level4:
rsMem.MoveFirst
For i = 0 To rsMem.RecordCount - 1 Step 1
    If Label3.Caption = rsMem!Idno Then
     Label4.Caption = rsMem!RefId
     no = no + 1
     GoTo Calculated
     End If
    rsMem.MoveNext
Next

Calculated:

'MsgBox no - 1

rsMem.MoveFirst
For i = 0 To rsMem.RecordCount - 1 Step 1
    If rsMem!Idno = Label1.Caption Then
       rsMem!Due = rsMem!Due + Level1
       rsMem!Balance = rsMem!Balance + Level1
       rsMem.Update
  
    End If
    rsMem.MoveNext
Next


rsMem.MoveFirst
For i = 0 To rsMem.RecordCount - 1 Step 1
    If rsMem!Idno = Label2.Caption Then
       rsMem!Due = rsMem!Due + Level2
       rsMem!Balance = rsMem!Balance + Level2
       rsMem.Update
      End If
    rsMem.MoveNext
Next


rsMem.MoveFirst
For i = 0 To rsMem.RecordCount - 1 Step 1
    If rsMem!Idno = Label3.Caption Then
       rsMem!Due = rsMem!Due + Level3
       rsMem!Balance = rsMem!Balance + Level3
       rsMem.Update
  
    End If
    rsMem.MoveNext
Next

rsMem.MoveFirst
For i = 0 To rsMem.RecordCount - 1 Step 1
    If rsMem!Idno = Label4.Caption Then
       rsMem!Due = rsMem!Due + Level4
       rsMem!Balance = rsMem!Balance + Level4
       rsMem.Update
  
    End If
    rsMem.MoveNext
Next


End Sub

Private Sub Command2_Click()

' cnt is a counter used to check that the commission is not added at a 0 level, but it is added from 1 to 4 Levels only
cnt = 0
Level = 0
lblID.Caption = Text1.Text

again:

If Level = 1 Then comm = Level1
If Level = 2 Then comm = Level2
If Level = 3 Then comm = Level3
If Level = 4 Then comm = Level4
If Level < 1 Or Level > 4 Then comm = 0

If rsMem.RecordCount > 0 Then rsMem.MoveFirst

For i = 0 To rsMem.RecordCount - 1 Step 1
    
    If Text1.Text = rsMem!Idno Then
    
If Level > 0 Then
Call AddAccount
'rsAcc.AddNew
'rsAcc!Id = rsMem!Idno
'rsAcc!Name = rsMem!Name
'rsAcc!Level = Level
'rsAcc!RefId = lblID.Caption
'rsAcc!RefName = "Pagal"
'rsAcc!com = comm
'rsAcc.Update
End If
    
    
       cnt = cnt + 1
       Level = Level + 1
    
       Text1.Text = rsMem!RefId
  '     MsgBox (Text1.Text)
       
'Call AddAccount

     If cnt > 1 Then
       rsMem!Due = rsMem!Due + comm
       rsMem!Balance = rsMem!Balance + comm
       rsMem.Update
     End If
     
    GoTo again
    End If
    rsMem.MoveNext
Next

End Sub

Private Sub AddAccount()
rsAcc.AddNew
rsAcc!Id = rsMem!Idno
rsAcc!Name = rsMem!Name
rsAcc!Level = Level
rsAcc!RefId = lblID.Caption
rsAcc!RefName = "Pagal"
rsAcc!com = comm
rsAcc.Update
End Sub


Private Sub Form_Load()

If rsComm.RecordCount > 0 Then rsComm.MoveFirst

Level1 = rsComm!comm    'Level1 = 100
If rsComm.EOF <> True Then rsComm.MoveNext
Level2 = rsComm!comm    'Level2 = 50
If rsComm.EOF <> True Then rsComm.MoveNext
Level3 = rsComm!comm    'Level3 = 25
If rsComm.EOF <> True Then rsComm.MoveNext
Level4 = rsComm!comm    'Level4 = 25

End Sub
