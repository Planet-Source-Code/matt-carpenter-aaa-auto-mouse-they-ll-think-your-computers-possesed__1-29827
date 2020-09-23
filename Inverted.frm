VERSION 5.00
Begin VB.Form Inverted 
   Caption         =   "Invert"
   ClientHeight    =   5610
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7815
   LinkTopic       =   "Form2"
   ScaleHeight     =   374
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   521
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "Inverted"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public pointX As Integer
Public pointY As Integer
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long

Private Sub Form_Click()
End
Me.Visible = False

End Sub

Private Sub Form_Load()
pointX = 100
pointY = 100
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Print MouseMove
Me.Visible = False


If x > pointX Then
  SetCursorPos x - 20, y
  pointX = x - 20
End If

If x < pointX Then
  SetCursorPos x + 20, y
  pointX = x + 20
End If

If y > pointY Then
  SetCursorPos x, y - 20
  pointY = y - 20
End If

If y < pointY Then
  SetCursorPos x, y + 20
  pointY = y + 20
End If

End Sub
