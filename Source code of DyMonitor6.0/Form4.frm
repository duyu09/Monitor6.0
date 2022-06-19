VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   0  'None
   Caption         =   "DUYU - REC"
   ClientHeight    =   690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1665
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   690
   ScaleWidth      =   1665
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "REC"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "¡¤"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   42
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   375
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Static ox As Integer, oy As Integer
  If Button = 1 Then
    Me.Left = Me.Left + X - ox
    Me.Top = Me.Top + Y - oy
  Else
    ox = X
    oy = Y
  End If
End Sub


Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Static ox As Integer, oy As Integer
  If Button = 1 Then
    Me.Left = Me.Left + X - ox
    Me.Top = Me.Top + Y - oy
  Else
    ox = X
    oy = Y
  End If
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Static ox As Integer, oy As Integer
  If Button = 1 Then
    Me.Left = Me.Left + X - ox
    Me.Top = Me.Top + Y - oy
  Else
    ox = X
    oy = Y
  End If
End Sub
