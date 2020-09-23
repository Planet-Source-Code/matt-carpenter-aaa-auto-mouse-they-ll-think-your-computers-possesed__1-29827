VERSION 5.00
Begin VB.Form frmRecord 
   Caption         =   "Mouser Record Window"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7395
   LinkTopic       =   "Form2"
   ScaleHeight     =   346
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   493
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer wait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1800
      Top             =   840
   End
   Begin VB.Timer Play 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1200
      Tag             =   "0"
      Top             =   240
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Record"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "FYI: When you test this, to get out, just do ctl+alt+delete, then end task..."
      Height          =   495
      Left            =   3600
      TabIndex        =   2
      Top             =   3840
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Click Record, then drag your mouse around this window in annoying patterns. Then click play..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3600
      TabIndex        =   1
      Top             =   3120
      Width           =   5895
   End
End
Attribute VB_Name = "frmRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public record As String
Public timetoplay As String
Public recorder As Boolean
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long

Private Sub Command1_Click()
If Command1.Caption = "Play" Then
Unload Form1

Dim howmany As Long


timetoplay = InputBox("Play when? [hh:mm]", "Mouser!")
wait = True



recorder = False

movements = Split(record, " ", -1, vbBinaryCompare)
Me.Visible = False

wait.Enabled = True

Exit Sub
End If


recorder = True
Command1.Caption = "Play"


End Sub

Private Sub Command2_Click()
Dim howmany As Long
saveas = InputBox("Record:", "Record", record)


howmany = InputBox("Play in how many seconds?")
howmany = howmany * 1000
PlayS.Interval = howmany


recorder = False
Command2.Enabled = False
movements = Split(record, " ", -1, vbBinaryCompare)
Me.Visible = False

PlayS.Enabled = True

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If recorder = True Then
  record = record & x & "," & y & " "
End If

End Sub

Private Sub Timer1_Timer()

End Sub

Private Sub Play_Timer()
On Error GoTo err
movements = Split(record, " ", -1, vbBinaryCompare)

aryXY = Split(movements(Play.Tag), ",", -1, vbBinaryCompare)
  SetCursorPos aryXY(0), aryXY(1)
Play.Tag = Play.Tag + 1
Exit Sub
err:
Play.Tag = 0

End Sub

Private Sub PlayS_Timer()
Play.Enabled = True
PlayS.Enabled = False

End Sub

Private Sub wait_Timer()
aryTime = Split(Time, ":", -1, vbBinaryCompare)
hh = aryTime(0)
mm = aryTime(1)
If Len(hh) = 1 Then hh = "0" & hh
aryDefinedTime = Split(timetoplay, ":", -1, vbBinaryCompare)
hh2 = aryDefinedTime(0)
mm2 = aryDefinedTime(1)
If hh = hh2 And mm = mm2 Then
'Hit!
Play.Enabled = True
wait.Enabled = False
End If



End Sub
