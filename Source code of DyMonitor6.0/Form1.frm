VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000014&
   Caption         =   "DuyuMonitor"
   ClientHeight    =   7815
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12000
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   12000
   StartUpPosition =   2  '��Ļ����
   Begin VB.Timer Timer6 
      Interval        =   2222
      Left            =   4560
      Top             =   6720
   End
   Begin VB.Timer Timer5 
      Interval        =   1
      Left            =   4080
      Top             =   6720
   End
   Begin VB.Frame Frame5 
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7815
      Left            =   8400
      TabIndex        =   12
      Top             =   0
      Width           =   3615
      Begin VB.CommandButton Command7 
         Caption         =   "About"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   37
         Top             =   7320
         Width           =   1095
      End
      Begin VB.CommandButton Command10 
         Caption         =   "��������ͷ"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   36
         Top             =   960
         Width           =   1455
      End
      Begin VB.CheckBox Check5 
         Caption         =   "ǿ��"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   32
         Top             =   6360
         Width           =   735
      End
      Begin VB.CheckBox Check4 
         Caption         =   "ǿ��ת�ƽ��㵽�������������ʹ�ã�"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   31
         Top             =   6840
         Width           =   2175
      End
      Begin VB.CheckBox Check3 
         Caption         =   "ǰ����ʾ���ö���"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   30
         Top             =   6360
         Width           =   1815
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Duyu���������"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         Picture         =   "Form1.frx":1856A
         TabIndex        =   29
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H000000FF&
         Caption         =   "����ǿ�ƹػ�"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         MaskColor       =   &H000000FF&
         TabIndex        =   28
         ToolTipText     =   "���棡������Ҫʹ�ô˹��ܣ��������ϵͳ������ʱ���ջ����壡"
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C0FF&
         Caption         =   "�ر�"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         MaskColor       =   &H000000FF&
         TabIndex        =   27
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "��¼������ϸ��Ϣ"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   26
         Top             =   5880
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox Check2 
         Caption         =   "����Ļ���Ͻ���ʾ¼��ͼ��"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   25
         Top             =   5400
         Width           =   2655
      End
      Begin VB.Frame Frame2 
         Caption         =   "��¼�����趨    "
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         TabIndex        =   20
         Top             =   4320
         Width           =   3135
         Begin VB.OptionButton Option1 
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton Option2 
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   960
            TabIndex        =   22
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox Text3 
            Height          =   270
            Left            =   1800
            TabIndex        =   21
            Text            =   "100"
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2640
            TabIndex        =   24
            Top             =   360
            Width           =   255
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "��Ϣ��ȫ   "
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         TabIndex        =   17
         Top             =   2880
         Width           =   3135
         Begin VB.CommandButton Command5 
            Caption         =   "����Ϣ�洢Ŀ¼����"
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton Command6 
            Caption         =   "����Ϣ�洢Ŀ¼����"
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   1680
            TabIndex        =   18
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "�����¼��   "
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         TabIndex        =   13
         Top             =   1680
         Width           =   3135
         Begin VB.CommandButton cmdRecord 
            Caption         =   "¼��"
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton cmdStop 
            Caption         =   "ֹͣ"
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1200
            TabIndex        =   15
            Top             =   360
            Width           =   855
         End
         Begin VB.CommandButton cmdPlay 
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2160
            TabIndex        =   14
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Label Label9 
         Caption         =   ")"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   34
         Top             =   6480
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "("
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   33
         Top             =   6480
         Width           =   375
      End
   End
   Begin VB.Timer Timer4 
      Interval        =   10
      Left            =   3120
      Top             =   6720
   End
   Begin VB.Timer Timer3 
      Interval        =   20
      Left            =   3600
      Top             =   6720
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2175
      Begin VB.CommandButton Command9 
         Caption         =   "ˢ��"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1170
         TabIndex        =   35
         Top             =   4320
         Width           =   735
      End
      Begin VB.Timer Timer2 
         Interval        =   20
         Left            =   0
         Top             =   2280
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2355
         Left            =   120
         TabIndex        =   8
         Top             =   4680
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   5
         Text            =   "10"
         Top             =   840
         Width           =   1455
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   0
         Top             =   3000
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H000000FF&
         Caption         =   "��ʼ���"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         MaskColor       =   &H000000FF&
         TabIndex        =   4
         Top             =   3360
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   1680
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   1
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   7200
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "ǰ̨���̣�"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1560
         TabIndex        =   7
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "��Ϣ�洢ʱ���� : "
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "��Ϣ�洢Ŀ¼: "
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   1335
      End
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "���ڼ����,������С����ť������С������."
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3240
      TabIndex        =   11
      Top             =   1920
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type GUID
Data1 As Long
Data2 As Integer
Data3 As Integer
Data4(0 To 7) As Byte
End Type
Private Type GdiplusStartupInput
GdiplusVersion As Long
DebugEventCallback As Long
SuppressBackgroundThread As Long
SuppressExternalCodecs As Long
End Type
Private Type EncoderParameter
GUID As GUID
NumberOfValues As Long
type As Long
Value As Long
End Type
Private Type EncoderParameters
Count As Long
Parameter As EncoderParameter
End Type
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Function GdiplusStartup Lib "GDIPlus" (token As Long, inputbuf As GdiplusStartupInput, ByVal outputbuf As Long) As Long
Private Declare Function GdiplusShutdown Lib "GDIPlus" (ByVal token As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "GDIPlus" (ByVal hbm As Long, ByVal hpal As Long, Bitmap As Long) As Long
Private Declare Function GdipDisposeImage Lib "GDIPlus" (ByVal Image As Long) As Long
Private Declare Function GdipSaveImageToFile Lib "GDIPlus" (ByVal Image As Long, ByVal filename As Long, clsidEncoder As GUID, encoderParams As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal str As Long, id As GUID) As Long
Private Declare Function GdipCreateBitmapFromFile Lib "GDIPlus" (ByVal filename As Long, Bitmap As Long) As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Declare Function RtlAdjustPrivilege& Lib "ntdll" (ByVal Privilege&, ByVal Newvalue&, ByVal NewThread&, Oldvalue&)
Private Declare Function NtShutdownSystem& Lib "ntdll" (ByVal ShutdownAction&)
Const SE_SHUTDOWN_PRIVILEGE& = 19
Const SHUTDOWN& = 0
Const RESTART& = 1
Const POWEROFF& = 2
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd _
As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const WM_SYSCOMMAND = &H112&
Const SC_MONITORPOWER = &HF170&


Private Declare Function GetDriveType Lib "kernel32" Alias _
        "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function GetVolumeInformation Lib "kernel32" Alias _
        "GetVolumeInformationA" (ByVal lpRootPathName As String, _
        ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize _
        As Long, lpVolumeSerialNumber As Long, _
        lpMaximumComponentLength As Long, lpFileSystemFlags As _
        Long, ByVal lpFileSystemNameBuffer As String, ByVal _
        nFileSystemNameSize As Long) As Long
Private Declare Function SetVolumeLabel Lib "kernel32" Alias _
        "SetVolumeLabelW" (ByVal lpRootPathName As String, ByVal _
        lpVolumeName As String) As Long

Option Explicit
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long

Dim ReturnString As String * 256
Dim RetValue As Long
Dim errorstring As String * 1024

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST& = -1
' �����������б�������λ���κ�������ڵ�ǰ��
Const HWND_NOTOPMOST = -2 'ȡ�����ϲ��趨
Private Const SWP_NOSIZE& = &H1
' ���ִ��ڴ�С
Private Const SWP_NOMOVE& = &H2
' ���ִ���λ��




Private Sub Check3_Click()
If Check3.Value = 1 Then
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
 If Check4.Value = 1 Then
 Form1.SetFocus
 End If
Else
SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End If
End Sub

Private Sub cmdPlay_Click()
Dim CommandString As String
If Dir(App.Path & "\Dy_Monitor.Wav") <> "" Then
RetValue = mciSendString("open Dy_Monitor.Wav alias sounds1", ReturnString, 256, 0)
mciSendString "play sounds1", ReturnString, 256, 0
Else
MsgBox "û�������ļ�", vbOKOnly, "��ʾ"
End If
End Sub
Private Sub cmdRecord_Click()
RetValue = mciSendString("open new type waveaudio alias sounds1", ReturnString, 256, 0)
If RetValue = 0 Then '��ʼ¼��...
RetValue = mciSendString("record sounds1", ReturnString, 256, 0)
cmdStop.Enabled = True
Else
mciGetErrorString RetValue, errorstring, 1024
MsgBox errorstring
End If
End Sub
Private Sub cmdStop_Click()
RetValue = mciSendString("stop sounds1", ReturnString, 256, 0)
RetValue = mciSendString("save sounds1 Dy_Monitor.Wav", ReturnString, 256, 0)
If RetValue <> 0 Then
mciGetErrorString RetValue, errorstring, 1024
MsgBox errorstring
End If
RetValue = mciSendString("close sounds1", ReturnString, 256, 0)
cmdStop.Enabled = False
On Error Resume Next
FileCopy App.Path & "\Dy_Monitor.Wav", Text1.Text
End Sub

Private Sub Command10_Click()
Form5.Show
Form6.Show
End Sub

Private Sub Command7_Click()
Form2.Show
End Sub

Private Sub Command8_Click()
Form3.Show
End Sub

Private Sub Command9_Click()
List1.Clear
EnumWindows AddressOf EnumWindowsProc, 0&
End Sub

Private Sub Form_Unload(Cancel As Integer)
mciSendString "close all", ReturnString, 256, 0
End
End Sub

Private Sub Command1_Click()
Dim ss As Integer
Dim sfile As String, ret As Integer
If Dir(Text1.Text, vbDirectory) = "" Then
ss = MsgBox("Ŀ¼�����ڣ�", 48)
Exit Sub
End If
List1.Clear
EnumWindows AddressOf EnumWindowsProc, 0&
Text1.Text = Text1.Text & "\"
    Form1.AutoRedraw = True
    BitBlt Form1.hDC, 0, 0, Screen.Width, Screen.Height, GetDC(0), 0, 0, vbSrcCopy
    On Error Resume Next
    sfile = Text1.Text & Replace(Time(), ":", "_") & ".temp.bmp"
       If Err.Number = 75 Then
       MsgBox ("�޷����ļ���Ŀ¼,��������Ϊû�н�����")
       Exit Sub
       End If
    SavePicture Form1.Image, sfile
    Kill Text1.Text & Replace(Time(), ":", "_") & ".temp.bmp"
    ret = PictureBoxSaveJPG(Form1.Image, CStr(Text1.Text & Replace(Time(), ":", "_") & ".jpg"))
If Check1.Value = 1 Then
Open Text1.Text & "Record_DyMonitor.txt" For Append As #1
  If Timer1.Tag = 1 Then
  Timer1.Tag = 0
  Print #1, "�ӿ�����" & Now() & "������" & GetTickCount() & "����"
  Print #1, "Լ��:" & Format(DateAdd("s", CDec(CLng(GetTickCount() / 1000)), "00:00:00"), "HH:mm:ss")
  Print #1, "--------------------------------------"
  End If
Print #1, Date & " " & Time() & "_ " & "�ѽ������档"
 If List1.ListCount > 0 Then
 Print #1, "---------------------------------------"
 Print #1, "ǰ̨����_������: "
 Dim der As Long
 For der = 1 To List1.ListCount
 Print #1, List1.List(der - 1)
 Next der
 End If

Dim rrrt As String
Dim ids As Long
    For ids = Asc("A") To Asc("Z")
        If GetDriveType(Chr(ids) + ":") = 2 Then
           Print #1, "---------------------------------"
           Print #1, "�ֲ���������U��: " & Chr(ids) & ":\��"
           rrrt = Chr(ids)
        End If
    Next ids

Dim sVolName As String * 256
Dim sFileSys As String * 256
Dim lVolSerial As Long                '������ֱ���
Dim lMC As Long
Dim lFileFlag As Long
            '����GetVolumeInformation�����������ѡ�̷��ľ�ꡢ������ʽ��Ϣ
GetVolumeInformation rrrt + ":\", sVolName, 256, lVolSerial, lMC, lFileFlag, sFileSys, 256
Print #1, "U�̵ľ��Ϊ��" & sVolName

Close
End If
Frame2.Tag = Val(Frame2.Tag) - 1
Text1.Text = Left(Text1.Text, Len(Text1.Text) - 1)
End Sub
Private Function PictureBoxSaveJPG(ByVal pict As StdPicture, ByVal filename As String, Optional ByVal quality As Byte = 80) As Boolean
Dim tSI As GdiplusStartupInput
Dim lRes As Long
Dim lGDIP As Long
Dim lBitmap As Long
'��ʼ�� GDI+
tSI.GdiplusVersion = 1
lRes = GdiplusStartup(lGDIP, tSI, 0)
If lRes = 0 Then
'�Ӿ������ GDI+ ͼ��
lRes = GdipCreateBitmapFromHBITMAP(pict.Handle, 0, lBitmap)
If lRes = 0 Then
Dim tJpgEncoder As GUID
Dim tParams As EncoderParameters
'��ʼ����������GUID��ʶe799bee5baa6e997aee7ad94e78988e69d8331333332626663
CLSIDFromString StrPtr("{557CF401-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
'���ý���������
tParams.Count = 1
With tParams.Parameter ' Quality
'�õ�Quality������GUID��ʶ
CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"), .GUID
.NumberOfValues = 1
.type = 4
.Value = VarPtr(quality)
End With
'����ͼ��
lRes = GdipSaveImageToFile(lBitmap, StrPtr(filename), tJpgEncoder, tParams)
'����GDI+ͼ��
GdipDisposeImage lBitmap
End If
'���� GDI+
GdiplusShutdown lGDIP
End If
If lRes Then
PictureBoxSaveJPG = False
Else
PictureBoxSaveJPG = True
End If
End Function

Private Sub Command2_Click()
Dim ss As Integer
If Dir(Text1.Text, vbDirectory) = "" Then
ss = MsgBox("Ŀ¼�����ڣ�", 48)
Exit Sub
End If
If Command2.Caption = "��ʼ���" Then
aa = CLng(Timer())
Timer1.Enabled = True
Command2.Caption = "�رռ��"
Frame2.Tag = Val(Text3.Text)
Else
Timer1.Enabled = False
Command2.Caption = "��ʼ���"
Frame2.Tag = Val(Text3.Text)
End If
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Command4_Click()
RtlAdjustPrivilege& SE_SHUTDOWN_PRIVILEGE&, 1, 0, 0 '����Ȩ��
SendMessage Me.hwnd, WM_SYSCOMMAND, SC_MONITORPOWER, ByVal 2& '�ر���ʾ��
NtShutdownSystem& SHUTDOWN& Or POWEROFF& '�ػ�
On Error Resume Next
Shell "cmd.exe /c shutdown -s -t 1"
End Sub

Private Sub Command5_Click()
Dim ss As Integer
If Dir(Text1.Text, vbDirectory) = "" Then
ss = MsgBox("Ŀ¼������", 48)
Exit Sub
End If
On Error Resume Next
Dim wsdf As String
wsdf = Text1.Text
Name Text1.Text As Text1.Text & ".{00021401-0000-0000-C000-000000000046}"
Text1.Text = Text1.Text & ".{00021401-0000-0000-C000-000000000046}"
Shell ("cacls.exe " & Chr(34) & Text1.Text & Chr(34) & " /e /t /p everyone:n")
Shell ("cacls.exe " & Chr(34) & Text1.Text & Chr(34) & " /e /t /p system:n")
Shell ("cacls.exe " & Chr(34) & Text1.Text & Chr(34) & " /e /t /p Administrators:n")
Shell ("cacls.exe " & Chr(34) & Text1.Text & Chr(34) & " /e /t /p Administrator:n")
Shell ("cacls.exe " & Chr(34) & Text1.Text & Chr(34) & " /e /t /p Authenticated Users:n")
Shell ("cacls.exe " & Chr(34) & Text1.Text & Chr(34) & " /e /t /p %username%:n")
Text1.Text = wsdf
MsgBox ("��ϡ�")
End Sub

Private Sub Command6_Click()
Dim ss As Integer
If Dir(Text1.Text & ".{00021401-0000-0000-C000-000000000046}", vbDirectory) = "" Then
ss = MsgBox("Ŀ¼������", 48)
Exit Sub
End If
On Error Resume Next
Dim wsdf As String
wsdf = Text1.Text
Text1.Text = Text1.Text & ".{00021401-0000-0000-C000-000000000046}"
Shell ("cacls.exe " & Chr(34) & Text1.Text & Chr(34) & " /e /t /p everyone:f")
Shell ("cacls.exe " & Chr(34) & Text1.Text & Chr(34) & " /e /t /p system:f")
Shell ("cacls.exe " & Chr(34) & Text1.Text & Chr(34) & " /e /t /p Administrators:f")
Shell ("cacls.exe " & Chr(34) & Text1.Text & Chr(34) & " /e /t /p Administrator:f")
Shell ("cacls.exe " & Chr(34) & Text1.Text & Chr(34) & " /e /t /p Authenticated Users:f")
Shell ("cacls.exe " & Chr(34) & Text1.Text & Chr(34) & " /e /t /p %username%:f")
Sleep (1111)
Name Text1.Text As wsdf
Text1.Text = wsdf
MsgBox ("��ϡ�")
End Sub

Private Sub Label5_Click()
Shell (Chr(34) & "C:\Windows\system32\rundll32.exe" & Chr(34) & " Shell32.dll,Control_RunDLL " & Chr(34) & "C:\Windows\system32\timedate.cpl" & Chr(34))
End Sub

Private Sub List1_dblClick()
Shell "taskmgr.exe", vbNormalFocus
End Sub

Private Sub Text3_Change()
Frame2.Tag = Val(Text3.Text)
End Sub

Private Sub Timer1_Timer()
If CLng(Timer()) = aa Then
aa = CLng(Timer() + 60 * Val(Text2.Text))
Command1_Click
End If
 If Option2.Value = True Then
  If Frame2.Tag = 0 And Command2.Caption <> "��ʼ���" Then
  Command2_Click
  End If
  End If
End Sub
Private Sub Form_Load()
Timer1.Tag = 1
 '���°ѳ������System Tray====================================System Tray Begin
 With nfIconData
 .hwnd = Me.hwnd
 .uID = Me.Icon
 .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
 .uCallbackMessage = WM_MOUSEMOVE
 .hIcon = Me.Icon.Handle
 '��������ƶ���������ʱ��ʾ��Tip
 .szTip = "DuyuMonitor" & vbNullChar
 .cbSize = Len(nfIconData)
 End With
 Call Shell_NotifyIcon(NIM_ADD, nfIconData)
 '=============================================================System Tray End
Text1.Text = App.Path
EnumWindows AddressOf EnumWindowsProc, 0&
SendMessage List1.hwnd, &H194, 3000, ByVal 0
 End Sub

'5����Form1��QueryUnload�¼���д�����´���:

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call Shell_NotifyIcon(NIM_DELETE, nfIconData)
End Sub

'6����Form1��MouseMove�¼���д�����´���:

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 Dim lMsg As Single
 lMsg = x / Screen.TwipsPerPixelX
 Select Case lMsg
 Case WM_LBUTTONUP
 'MsgBox "��������Ҽ����ͼ��!", vbInformation,
 '�����������ʾ����
 ShowWindow Me.hwnd, SW_RESTORE
 '���������Ŀ���ǰѴ�����ʾ�ڴ������
 'Me.Show
 'Me.SetFocus
 '' Case WM_RBUTTONUP
 '' PopupMenu MenuTray '�������ϵͳTrayͼ���ϵ��Ҽ����򵯳��˵�MenuTray
 '' Case WM_MOUSEMOVE
 '' Case WM_LBUTTONDOWN
 '' Case WM_LBUTTONDBLCLK
 '' Case WM_RBUTTONDOWN
 '' Case WM_RBUTTONDBLCLK
 '' Case Else
 End Select
End Sub


Private Sub Form_Resize()
If Form1.WindowState = vbMinimized Then
Me.Hide
End If
End Sub

Private Sub Timer2_Timer()
Label5.Caption = Time()
End Sub

Private Sub Timer3_Timer()
If Command2.Caption = "��ʼ���" Then
Label7.Visible = False
Else
Label7.Visible = True
End If
End Sub

Private Sub Timer4_Timer()
If Check2.Value = 1 Then
 If Command2.Caption = "�رռ��" And Form4.Visible = False Then
 Form4.Visible = True
 End If
Else
Form4.Visible = False
End If
End Sub

Private Sub Timer5_Timer()
If Check3.Value = 1 And Check5.Value = 1 Then
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End If
End Sub

Private Sub Timer6_Timer()
List1.Clear
EnumWindows AddressOf EnumWindowsProc, 0&
End Sub
