VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   2835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4650
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2835
   ScaleWidth      =   4650
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   2160
      Top             =   720
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enough reading, let's go!"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "I'm not responsible if you plant it on your bosses computer and you get fired..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   2400
      Width           =   4695
   End
   Begin VB.Label Label2 
      Caption         =   $"frmMain.frx":0442
      Height          =   975
      Left            =   720
      TabIndex        =   2
      Top             =   1200
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   $"frmMain.frx":04E4
      Height          =   975
      Left            =   720
      TabIndex        =   1
      Top             =   240
      Width           =   3855
   End
   Begin VB.Image mouse 
      Height          =   480
      Left            =   120
      Picture         =   "frmMain.frx":058C
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public mm As Integer
Public captionn As Integer


Private Sub Command2_Click()
Inverted.Show



End Sub

Private Sub Command1_Click()
frmRecord.Show

 
End Sub

Private Sub RandomPoints_Timer()
Randomize
xx = Int(Rnd * 1024)
yy = Int(Rnd * 768)
SetCursorPos xx, yy

End Sub

Private Sub run_Timer()
 If random = True Then
    'Random Cursor Points Now
    Me.Visible = False
    
    RandomPoints.Enabled = True
  End If
  If record = True Then
    'record and play now
    frmRecord.Show
    Me.Visible = False
  End If
  run.Enabled = False
  
End Sub

Private Sub Form_Load()
mm = 100
captionn = 12

End Sub

Private Sub Timer1_Timer()
mouse.Top = mouse.Top + mm
If mouse.Top > 2000 Then
mm = -100
End If
If mouse.Top < 1 Then
mm = 100
End If


End Sub

Private Sub Timer2_Timer()
Me.caption = Me.caption & Left(Right("Mouser 2.0  ", captionn), 1)
If Len(Me.caption) = 11 Then
captionn = 12
Me.caption = ""
Exit Sub
End If
captionn = captionn - 1


End Sub
