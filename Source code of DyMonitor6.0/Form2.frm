VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����-���Լ������DuyuMonitor�� ��Ȩ����"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7005
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   7005
   StartUpPosition =   2  '��Ļ����
   Begin VB.Label Label4 
      Caption         =   "ѧУ������https://www.lcez.cn/"
      Height          =   255
      Left            =   4200
      TabIndex        =   3
      ToolTipText     =   "�����������Ƕ��й���"
      Top             =   2640
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "DuyuMonitor"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "��������Ȩ��"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   1
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "���������ǵڶ���ѧ55��31�� ����NO.028"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2520
      TabIndex        =   0
      Top             =   600
      Width           =   4215
   End
   Begin VB.Image Image1 
      Height          =   1920
      Left            =   360
      Picture         =   "Form2.frx":1856A
      Top             =   360
      Width           =   1920
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label4_Click()
Shell ("explorer.exe https://www.lcez.cn/"), vbNormalFocus
End Sub
