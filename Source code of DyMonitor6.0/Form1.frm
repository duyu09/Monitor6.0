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
   StartUpPosition =   2  '屏幕中心
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
         Name            =   "微软雅黑"
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
            Name            =   "黑体"
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
         Caption         =   "调用摄像头"
         BeginProperty Font 
            Name            =   "微软雅黑"
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
         Caption         =   "强制"
         BeginProperty Font 
            Name            =   "微软雅黑"
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
         Caption         =   "强制转移焦点到此软件（不建议使用）"
         BeginProperty Font 
            Name            =   "微软雅黑"
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
         Caption         =   "前端显示（置顶）"
         BeginProperty Font 
            Name            =   "微软雅黑"
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
         Caption         =   "Duyu任务管理器"
         BeginProperty Font 
            Name            =   "微软雅黑"
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
         Caption         =   "快速强制关机"
         BeginProperty Font 
            Name            =   "微软雅黑"
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
         ToolTipText     =   "警告！尽量不要使用此功能！后果是损害系统，严重时会烧坏主板！"
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C0FF&
         Caption         =   "关闭"
         BeginProperty Font 
            Name            =   "微软雅黑"
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
         Caption         =   "记录运行详细信息"
         BeginProperty Font 
            Name            =   "微软雅黑"
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
         Caption         =   "在屏幕左上角显示录像图标"
         BeginProperty Font 
            Name            =   "微软雅黑"
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
         Caption         =   "记录次数设定    "
         BeginProperty Font 
            Name            =   "微软雅黑"
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
            Caption         =   "不限"
            BeginProperty Font 
               Name            =   "微软雅黑"
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
            Caption         =   "限制"
            BeginProperty Font 
               Name            =   "微软雅黑"
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
            Caption         =   "次"
            BeginProperty Font 
               Name            =   "微软雅黑"
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
         Caption         =   "信息安全   "
         BeginProperty Font 
            Name            =   "微软雅黑"
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
            Caption         =   "对信息存储目录上锁"
            BeginProperty Font 
               Name            =   "微软雅黑"
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
            Caption         =   "对信息存储目录解锁"
            BeginProperty Font 
               Name            =   "微软雅黑"
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
         Caption         =   "计算机录音   "
         BeginProperty Font 
            Name            =   "微软雅黑"
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
            Caption         =   "录音"
            BeginProperty Font 
               Name            =   "微软雅黑"
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
            Caption         =   "停止"
            BeginProperty Font 
               Name            =   "微软雅黑"
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
            Caption         =   "播放"
            BeginProperty Font 
               Name            =   "微软雅黑"
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
            Name            =   "宋体"
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
            Name            =   "宋体"
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
         Name            =   "微软雅黑"
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
         Caption         =   "刷新"
         BeginProperty Font 
            Name            =   "微软雅黑"
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
            Name            =   "微软雅黑"
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
            Name            =   "微软雅黑"
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
         Caption         =   "开始监控"
         BeginProperty Font 
            Name            =   "微软雅黑"
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
            Name            =   "微软雅黑"
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
         Caption         =   "立即保存"
         BeginProperty Font 
            Name            =   "微软雅黑"
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
            Name            =   "微软雅黑"
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
         Caption         =   "前台进程："
         BeginProperty Font 
            Name            =   "微软雅黑"
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
         Caption         =   "分钟"
         BeginProperty Font 
            Name            =   "微软雅黑"
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
         Caption         =   "信息存储时间间隔 : "
         BeginProperty Font 
            Name            =   "微软雅黑"
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
         Caption         =   "信息存储目录: "
         BeginProperty Font 
            Name            =   "微软雅黑"
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
      Caption         =   "正在监控中,单击最小化按钮可以缩小到托盘."
      BeginProperty Font 
         Name            =   "微软雅黑"
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
' 将窗口置于列表顶部，并位于任何最顶部窗口的前面
Const HWND_NOTOPMOST = -2 '取消最上层设定
Private Const SWP_NOSIZE& = &H1
' 保持窗口大小
Private Const SWP_NOMOVE& = &H2
' 保持窗口位置




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
MsgBox "没有声音文件", vbOKOnly, "提示"
End If
End Sub
Private Sub cmdRecord_Click()
RetValue = mciSendString("open new type waveaudio alias sounds1", ReturnString, 256, 0)
If RetValue = 0 Then '开始录音...
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
ss = MsgBox("目录不存在！", 48)
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
       MsgBox ("无法打开文件或目录,可能是因为没有解锁。")
       Exit Sub
       End If
    SavePicture Form1.Image, sfile
    Kill Text1.Text & Replace(Time(), ":", "_") & ".temp.bmp"
    ret = PictureBoxSaveJPG(Form1.Image, CStr(Text1.Text & Replace(Time(), ":", "_") & ".jpg"))
If Check1.Value = 1 Then
Open Text1.Text & "Record_DyMonitor.txt" For Append As #1
  If Timer1.Tag = 1 Then
  Timer1.Tag = 0
  Print #1, "从开机到" & Now() & "所经历" & GetTickCount() & "毫秒"
  Print #1, "约合:" & Format(DateAdd("s", CDec(CLng(GetTickCount() / 1000)), "00:00:00"), "HH:mm:ss")
  Print #1, "--------------------------------------"
  End If
Print #1, Date & " " & Time() & "_ " & "已截屏保存。"
 If List1.ListCount > 0 Then
 Print #1, "---------------------------------------"
 Print #1, "前台进程_窗体名: "
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
           Print #1, "现插入计算机的U盘: " & Chr(ids) & ":\盘"
           rrrt = Chr(ids)
        End If
    Next ids

Dim sVolName As String * 256
Dim sFileSys As String * 256
Dim lVolSerial As Long                '定义各种变量
Dim lMC As Long
Dim lFileFlag As Long
            '调用GetVolumeInformation函数，获得所选盘符的卷标、分区格式信息
GetVolumeInformation rrrt + ":\", sVolName, 256, lVolSerial, lMC, lFileFlag, sFileSys, 256
Print #1, "U盘的卷标为：" & sVolName

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
'初始化 GDI+
tSI.GdiplusVersion = 1
lRes = GdiplusStartup(lGDIP, tSI, 0)
If lRes = 0 Then
'从句柄创建 GDI+ 图像
lRes = GdipCreateBitmapFromHBITMAP(pict.Handle, 0, lBitmap)
If lRes = 0 Then
Dim tJpgEncoder As GUID
Dim tParams As EncoderParameters
'初始化解码器的GUID标识e799bee5baa6e997aee7ad94e78988e69d8331333332626663
CLSIDFromString StrPtr("{557CF401-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
'设置解码器参数
tParams.Count = 1
With tParams.Parameter ' Quality
'得到Quality参数的GUID标识
CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"), .GUID
.NumberOfValues = 1
.type = 4
.Value = VarPtr(quality)
End With
'保存图像
lRes = GdipSaveImageToFile(lBitmap, StrPtr(filename), tJpgEncoder, tParams)
'销毁GDI+图像
GdipDisposeImage lBitmap
End If
'销毁 GDI+
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
ss = MsgBox("目录不存在！", 48)
Exit Sub
End If
If Command2.Caption = "开始监控" Then
aa = CLng(Timer())
Timer1.Enabled = True
Command2.Caption = "关闭监控"
Frame2.Tag = Val(Text3.Text)
Else
Timer1.Enabled = False
Command2.Caption = "开始监控"
Frame2.Tag = Val(Text3.Text)
End If
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Command4_Click()
RtlAdjustPrivilege& SE_SHUTDOWN_PRIVILEGE&, 1, 0, 0 '提升权限
SendMessage Me.hwnd, WM_SYSCOMMAND, SC_MONITORPOWER, ByVal 2& '关闭显示器
NtShutdownSystem& SHUTDOWN& Or POWEROFF& '关机
On Error Resume Next
Shell "cmd.exe /c shutdown -s -t 1"
End Sub

Private Sub Command5_Click()
Dim ss As Integer
If Dir(Text1.Text, vbDirectory) = "" Then
ss = MsgBox("目录不存在", 48)
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
MsgBox ("完毕。")
End Sub

Private Sub Command6_Click()
Dim ss As Integer
If Dir(Text1.Text & ".{00021401-0000-0000-C000-000000000046}", vbDirectory) = "" Then
ss = MsgBox("目录不存在", 48)
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
MsgBox ("完毕。")
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
  If Frame2.Tag = 0 And Command2.Caption <> "开始监控" Then
  Command2_Click
  End If
  End If
End Sub
Private Sub Form_Load()
Timer1.Tag = 1
 '以下把程序放入System Tray====================================System Tray Begin
 With nfIconData
 .hwnd = Me.hwnd
 .uID = Me.Icon
 .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
 .uCallbackMessage = WM_MOUSEMOVE
 .hIcon = Me.Icon.Handle
 '定义鼠标移动到托盘上时显示的Tip
 .szTip = "DuyuMonitor" & vbNullChar
 .cbSize = Len(nfIconData)
 End With
 Call Shell_NotifyIcon(NIM_ADD, nfIconData)
 '=============================================================System Tray End
Text1.Text = App.Path
EnumWindows AddressOf EnumWindowsProc, 0&
SendMessage List1.hwnd, &H194, 3000, ByVal 0
 End Sub

'5、在Form1的QueryUnload事件中写入如下代码:

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call Shell_NotifyIcon(NIM_DELETE, nfIconData)
End Sub

'6、在Form1的MouseMove事件中写下如下代码:

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 Dim lMsg As Single
 lMsg = x / Screen.TwipsPerPixelX
 Select Case lMsg
 Case WM_LBUTTONUP
 'MsgBox "请用鼠标右键点击图标!", vbInformation,
 '单击左键，显示窗体
 ShowWindow Me.hwnd, SW_RESTORE
 '下面两句的目的是把窗口显示在窗口最顶层
 'Me.Show
 'Me.SetFocus
 '' Case WM_RBUTTONUP
 '' PopupMenu MenuTray '如果是在系统Tray图标上点右键，则弹出菜单MenuTray
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
If Command2.Caption = "开始监控" Then
Label7.Visible = False
Else
Label7.Visible = True
End If
End Sub

Private Sub Timer4_Timer()
If Check2.Value = 1 Then
 If Command2.Caption = "关闭监控" And Form4.Visible = False Then
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
