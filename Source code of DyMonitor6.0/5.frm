VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "USB摄像头视频图像的监控、截图、录像 - DuyuMonitor - 引用 新兴网络  http://www.newxing.com "
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10575
   Icon            =   "5.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   10575
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command5 
      BackColor       =   &H000000FF&
      Caption         =   "视频捕捉"
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
      Left            =   9360
      MaskColor       =   &H000000FF&
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
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
      Left            =   3240
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
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
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
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
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
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
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "USB摄像头视频图像的监控、截图、录像 - DuyuMonitor - 引用 新兴网络  http://www.newxing.com "
      BeginProperty Font 
         Name            =   "华文仿宋"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   4
      Top             =   0
      Width           =   5175
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal nID As Long) As Long
Dim ctCapWin As Long, ctAviPath As String, ctPicPath As String, ctConnect As Boolean
'视频窗口控制消息常数
Const WS_CHILD = &H40000000: Const WS_VISIBLE = &H10000000
Const WS_Caption = &HC00000: Const WS_ThickFrame = &H40000
Const WM_USER = &H400                       '用户消息开始号
Const WM_CAP_Connect = WM_USER + 10         '连接一个摄像头
Const WM_CAP_DisConnect = WM_USER + 11      '断开一个摄像头的连接
Const WM_CAP_SET_PREVIEW = WM_USER + 50     '使预览模式有效或者失效
Const WM_CAP_SET_OVERLAY = WM_USER + 51     '使窗口处于叠加模式，也会自动地使预览模式失效。
Const WM_CAP_SET_PREVIEWRATE = WM_USER + 52 '设置在预览模式下帧的显示频率
Const WM_CAP_EDIT_COPY = WM_USER + 30       '将当前图像复制到剪贴板
Const WM_CAP_SEQUENCE = WM_USER + 62        '开始录像，录像未结束前不会返回。
Const WM_Cap_File_Set_File = WM_USER + 20   '设置当前的视频捕捉文件
Const WM_Cap_File_Get_File = WM_USER + 21   '得到当前的视频捕捉文件
Private Function CutPathFile(nStr As String, nPath As String, nFile As String)
   '分解出文件和目录
    Dim i As Long, s As Long
   
    For i = 1 To Len(nStr)
       If Mid(nStr, i, 1) = "\" Then s = i  '查找最后一个目录分隔符
    Next
    If s > 0 Then
       nPath = Left(nStr, s): nFile = Mid(nStr, s + 1)
    Else
       nPath = "": nFile = nStr
    End If
End Function


Private Sub Command1_Click()
    '创建视频窗口和连接摄像头
     Dim nStyle As Long, T As Long
    
     If ctCapWin = 0 Then '创建一个视频窗口，大小：640*480
         T = Me.ScaleY(Command1.Top + Command1.Height * 1.1, Me.ScaleMode, 3) '视频窗口垂直位置：像素
        'nStyle = WS_Child + WS_Visible + WS_Caption + WS_ThickFrame '子窗口(在Form1内)+可见+标题栏+边框
         nStyle = WS_CHILD + WS_VISIBLE '视频窗口无标题栏和边框
        'nStyle = WS_Visible '视频窗口为独立窗口，关闭主窗口视频窗口也会自动关闭
         ctCapWin = capCreateCaptureWindow("DuyuMonitor - 视频窗口", nStyle, 0, T, 640, 480, Me.hwnd, 0)
     End If
    
    '将视频窗口连接到摄像头，如无后面两条语句视频窗口画面不会变化
     SendMessage ctCapWin, WM_CAP_Connect, 0, 0          '连接摄像头
     SendMessage ctCapWin, WM_CAP_SET_PREVIEW, 1, 0      '第三个参数：1-预览模式有效,0-预览模式无效
     SendMessage ctCapWin, WM_CAP_SET_PREVIEWRATE, 30, 0 '第三个参数：设置预览显示频率为每秒 30 帧
     ctConnect = True: KjEnabled True
    '"请检检查摄像头连接，并确定没有其他用户和程序使用。"
End Sub

Private Sub Command2_Click()
     SendMessage ctCapWin, WM_CAP_DisConnect, 0, 0  '断开摄像头连接
     ctConnect = False: KjEnabled True
End Sub

Private Sub Command3_Click()
   '截图,保存为图片文件
     Dim f As String, s As Long, nPath As String, nStr As String
    
     nPath = Trim(ctPicPath)
     If nPath = "" Then nPath = App.Path & "\MyPic"
     If Right(nPath, 1) <> "\" Then nPath = nPath & "\"
    
     On Error Resume Next
     Do
        s = s + 1
        f = nPath & "MyPic-" & s & ".bmp"
        If Dir(f, 23) = "" Then Exit Do
     Loop
     On Error GoTo 0
    
     nStr = Trim(InputBox("设置图片保存的文件名:", "保存图片", f))
     If nStr = "" Then Exit Sub
     Call CutPathFile(nStr, nPath, f)  '分解出文件和目录
     
       MsgBox "在指定的位置无法建立目录：" & vbCrLf & nPath, vbInformation, "保存文件"

     ctPicPath = nPath: f = nPath & f
     If Dir(f, 23) <> "" Then
        If vbCancel = MsgBox("文件已存在，覆盖此文件吗？" & vbCrLf & f, vbInformation + vbOKCancel, "截图 - 文件覆盖") Then Exit Sub
        On Error GoTo Cuo
        SetAttr f, 0
        Kill f
        On Error GoTo 0
     End If
   
     
     Clipboard.Clear: SendMessage ctCapWin, WM_CAP_EDIT_COPY, 0, 0 '将当前图像复制到剪贴板
     On Error GoTo Cuo
     SavePicture Clipboard.GetData, f '保存为 Bmp 图像，要保存为 jpg 格式，参见： 将图片保存或转变为JPG格式
     Exit Sub
Cuo:
     MsgBox "无法写文件：" & vbCrLf & f & vbCrLf & "如果您多次尝试仍然出现错误，请单击 视频捕捉 换用一个组件.", vbInformation, "保存文件"
End Sub

Private Sub Command4_Click()
   '用摄像头录像，并保存为视频文件
   '如果不设置文件路径和名称，或路径不存在，视频窗口会使用默认文件名 C:\CAPTURE.AVI
     Dim f As String, s As Long, nPath As String, nStr As String
    
     nPath = Trim(ctAviPath)
     If nPath = "" Then nPath = App.Path & "\MyVideo"
     If Right(nPath, 1) <> "\" Then nPath = nPath & "\"
    
     On Error Resume Next
     Do
        s = s + 1
        f = nPath & "MyVideo-" & s & ".avi"
        If Dir(f, 23) = "" Then Exit Do
     Loop
     On Error GoTo 0
    
     nStr = Trim(InputBox("设置录像保存的文件名:", "录像保存的文件名", f))
     If nStr = "" Then Exit Sub
     Call CutPathFile(nStr, nPath, f)  '分解出文件和目录
     
        MsgBox "在指定的位置无法建立目录：" & vbCrLf & nPath, vbInformation, "保存文件"

     ctAviPath = nPath: f = nPath & f
     If Dir(f, 23) <> "" Then
        If vbCancel = MsgBox("文件已存在，覆盖此文件吗？" & vbCrLf & f, vbInformation + vbOKCancel, "视频 - 文件覆盖") Then Exit Sub
        On Error GoTo Cuo
        SetAttr f, 0
        Kill f
        On Error GoTo 0
     End If
    
     Me.Caption = "摄像头控制 - 正在录像（任意位置单击鼠标停止）": KjEnabled False: DoEvents
     SendMessage ctCapWin, WM_Cap_File_Set_File, 0, ByVal f '设置录像保存的文件
     SendMessage ctCapWin, WM_CAP_SEQUENCE, 0, 0            '开始录像。录像未结束前不会返回
     Me.Caption = "摄像头控制": KjEnabled True
   
     Exit Sub
Cuo:
     MsgBox "无法写文件：" & vbCrLf & f & vbCrLf & "如果您多次尝试仍然出现错误，请单击 视频捕捉 换用一个组件.", vbInformation, "保存文件"
End Sub

Private Sub Command5_Click()
frmMain.Show
Form5.Hide
Form6.Hide
End Sub

Private Sub Form_Load()
  '设置按钮及位置，实际可以在控件设计期间完成
    Dim H1 As Long
    Me.Caption = "摄像头控制"
    Command1.Caption = "连接": Command1.ToolTipText = "连接摄像头"
    Command2.Caption = "断开": Command2.ToolTipText = "断开与摄像头的连接"
    Command3.Caption = "截图": Command3.ToolTipText = "将当前图像保存为图片文件"
    Command4.Caption = "录像": Command4.ToolTipText = "开始录像，保存为视频文件"

    H1 = Me.TextHeight("A")
    Command1.Move H1 * 0.5, H1 * 0.5, H1 * 4, H1 * 2
    Command2.Move H1 * 5, H1 * 0.5, H1 * 4, H1 * 2
    Command3.Move H1 * 10, H1 * 0.5, H1 * 4, H1 * 2
    Command4.Move H1 * 15, H1 * 0.5, H1 * 4, H1 * 2
   '读出用户设置
    Call ReadSaveSet
    KjEnabled True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call ReadSaveSet(True) '保存用户设置
End Sub

Private Sub KjEnabled(nEnabled As Boolean)
    If nEnabled Then
       Command1.Enabled = Not ctConnect: Command2.Enabled = ctConnect
       Command3.Enabled = ctConnect: Command4.Enabled = ctConnect
    Else
       Command1.Enabled = nEnabled: Command2.Enabled = nEnabled
       Command3.Enabled = nEnabled: Command4.Enabled = nEnabled
    End If
End Sub
Private Sub ReadSaveSet(Optional IsSave As Boolean)
   '保存或读出用户设置的图片和视频默认保存目录
    Dim nKey As String, nSub As String
    nKey = "摄像头控制程序": nSub = "UserOpt"
    If IsSave Then
       SaveSetting nKey, nSub, "AviPath", ctAviPath
       SaveSetting nKey, nSub, "PicPath", ctPicPath
    Else
       ctAviPath = GetSetting(nKey, nSub, "AviPath", "")
       ctPicPath = GetSetting(nKey, nSub, "PicPath", "")
    End If
End Sub
