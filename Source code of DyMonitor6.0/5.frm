VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "USB����ͷ��Ƶͼ��ļ�ء���ͼ��¼�� - DuyuMonitor - ���� ��������  http://www.newxing.com "
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10575
   Icon            =   "5.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   10575
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command5 
      BackColor       =   &H000000FF&
      Caption         =   "��Ƶ��׽"
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
      Left            =   9360
      MaskColor       =   &H000000FF&
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
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
      Left            =   3240
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
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
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
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
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
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
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "USB����ͷ��Ƶͼ��ļ�ء���ͼ��¼�� - DuyuMonitor - ���� ��������  http://www.newxing.com "
      BeginProperty Font 
         Name            =   "���ķ���"
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
'��Ƶ���ڿ�����Ϣ����
Const WS_CHILD = &H40000000: Const WS_VISIBLE = &H10000000
Const WS_Caption = &HC00000: Const WS_ThickFrame = &H40000
Const WM_USER = &H400                       '�û���Ϣ��ʼ��
Const WM_CAP_Connect = WM_USER + 10         '����һ������ͷ
Const WM_CAP_DisConnect = WM_USER + 11      '�Ͽ�һ������ͷ������
Const WM_CAP_SET_PREVIEW = WM_USER + 50     'ʹԤ��ģʽ��Ч����ʧЧ
Const WM_CAP_SET_OVERLAY = WM_USER + 51     'ʹ���ڴ��ڵ���ģʽ��Ҳ���Զ���ʹԤ��ģʽʧЧ��
Const WM_CAP_SET_PREVIEWRATE = WM_USER + 52 '������Ԥ��ģʽ��֡����ʾƵ��
Const WM_CAP_EDIT_COPY = WM_USER + 30       '����ǰͼ���Ƶ�������
Const WM_CAP_SEQUENCE = WM_USER + 62        '��ʼ¼��¼��δ����ǰ���᷵�ء�
Const WM_Cap_File_Set_File = WM_USER + 20   '���õ�ǰ����Ƶ��׽�ļ�
Const WM_Cap_File_Get_File = WM_USER + 21   '�õ���ǰ����Ƶ��׽�ļ�
Private Function CutPathFile(nStr As String, nPath As String, nFile As String)
   '�ֽ���ļ���Ŀ¼
    Dim i As Long, s As Long
   
    For i = 1 To Len(nStr)
       If Mid(nStr, i, 1) = "\" Then s = i  '�������һ��Ŀ¼�ָ���
    Next
    If s > 0 Then
       nPath = Left(nStr, s): nFile = Mid(nStr, s + 1)
    Else
       nPath = "": nFile = nStr
    End If
End Function


Private Sub Command1_Click()
    '������Ƶ���ں���������ͷ
     Dim nStyle As Long, T As Long
    
     If ctCapWin = 0 Then '����һ����Ƶ���ڣ���С��640*480
         T = Me.ScaleY(Command1.Top + Command1.Height * 1.1, Me.ScaleMode, 3) '��Ƶ���ڴ�ֱλ�ã�����
        'nStyle = WS_Child + WS_Visible + WS_Caption + WS_ThickFrame '�Ӵ���(��Form1��)+�ɼ�+������+�߿�
         nStyle = WS_CHILD + WS_VISIBLE '��Ƶ�����ޱ������ͱ߿�
        'nStyle = WS_Visible '��Ƶ����Ϊ�������ڣ��ر���������Ƶ����Ҳ���Զ��ر�
         ctCapWin = capCreateCaptureWindow("DuyuMonitor - ��Ƶ����", nStyle, 0, T, 640, 480, Me.hwnd, 0)
     End If
    
    '����Ƶ�������ӵ�����ͷ�����޺������������Ƶ���ڻ��治��仯
     SendMessage ctCapWin, WM_CAP_Connect, 0, 0          '��������ͷ
     SendMessage ctCapWin, WM_CAP_SET_PREVIEW, 1, 0      '������������1-Ԥ��ģʽ��Ч,0-Ԥ��ģʽ��Ч
     SendMessage ctCapWin, WM_CAP_SET_PREVIEWRATE, 30, 0 '����������������Ԥ����ʾƵ��Ϊÿ�� 30 ֡
     ctConnect = True: KjEnabled True
    '"���������ͷ���ӣ���ȷ��û�������û��ͳ���ʹ�á�"
End Sub

Private Sub Command2_Click()
     SendMessage ctCapWin, WM_CAP_DisConnect, 0, 0  '�Ͽ�����ͷ����
     ctConnect = False: KjEnabled True
End Sub

Private Sub Command3_Click()
   '��ͼ,����ΪͼƬ�ļ�
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
    
     nStr = Trim(InputBox("����ͼƬ������ļ���:", "����ͼƬ", f))
     If nStr = "" Then Exit Sub
     Call CutPathFile(nStr, nPath, f)  '�ֽ���ļ���Ŀ¼
     
       MsgBox "��ָ����λ���޷�����Ŀ¼��" & vbCrLf & nPath, vbInformation, "�����ļ�"

     ctPicPath = nPath: f = nPath & f
     If Dir(f, 23) <> "" Then
        If vbCancel = MsgBox("�ļ��Ѵ��ڣ����Ǵ��ļ���" & vbCrLf & f, vbInformation + vbOKCancel, "��ͼ - �ļ�����") Then Exit Sub
        On Error GoTo Cuo
        SetAttr f, 0
        Kill f
        On Error GoTo 0
     End If
   
     
     Clipboard.Clear: SendMessage ctCapWin, WM_CAP_EDIT_COPY, 0, 0 '����ǰͼ���Ƶ�������
     On Error GoTo Cuo
     SavePicture Clipboard.GetData, f '����Ϊ Bmp ͼ��Ҫ����Ϊ jpg ��ʽ���μ��� ��ͼƬ�����ת��ΪJPG��ʽ
     Exit Sub
Cuo:
     MsgBox "�޷�д�ļ���" & vbCrLf & f & vbCrLf & "�������γ�����Ȼ���ִ����뵥�� ��Ƶ��׽ ����һ�����.", vbInformation, "�����ļ�"
End Sub

Private Sub Command4_Click()
   '������ͷ¼�񣬲�����Ϊ��Ƶ�ļ�
   '����������ļ�·�������ƣ���·�������ڣ���Ƶ���ڻ�ʹ��Ĭ���ļ��� C:\CAPTURE.AVI
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
    
     nStr = Trim(InputBox("����¼�񱣴���ļ���:", "¼�񱣴���ļ���", f))
     If nStr = "" Then Exit Sub
     Call CutPathFile(nStr, nPath, f)  '�ֽ���ļ���Ŀ¼
     
        MsgBox "��ָ����λ���޷�����Ŀ¼��" & vbCrLf & nPath, vbInformation, "�����ļ�"

     ctAviPath = nPath: f = nPath & f
     If Dir(f, 23) <> "" Then
        If vbCancel = MsgBox("�ļ��Ѵ��ڣ����Ǵ��ļ���" & vbCrLf & f, vbInformation + vbOKCancel, "��Ƶ - �ļ�����") Then Exit Sub
        On Error GoTo Cuo
        SetAttr f, 0
        Kill f
        On Error GoTo 0
     End If
    
     Me.Caption = "����ͷ���� - ����¼������λ�õ������ֹͣ��": KjEnabled False: DoEvents
     SendMessage ctCapWin, WM_Cap_File_Set_File, 0, ByVal f '����¼�񱣴���ļ�
     SendMessage ctCapWin, WM_CAP_SEQUENCE, 0, 0            '��ʼ¼��¼��δ����ǰ���᷵��
     Me.Caption = "����ͷ����": KjEnabled True
   
     Exit Sub
Cuo:
     MsgBox "�޷�д�ļ���" & vbCrLf & f & vbCrLf & "�������γ�����Ȼ���ִ����뵥�� ��Ƶ��׽ ����һ�����.", vbInformation, "�����ļ�"
End Sub

Private Sub Command5_Click()
frmMain.Show
Form5.Hide
Form6.Hide
End Sub

Private Sub Form_Load()
  '���ð�ť��λ�ã�ʵ�ʿ����ڿؼ�����ڼ����
    Dim H1 As Long
    Me.Caption = "����ͷ����"
    Command1.Caption = "����": Command1.ToolTipText = "��������ͷ"
    Command2.Caption = "�Ͽ�": Command2.ToolTipText = "�Ͽ�������ͷ������"
    Command3.Caption = "��ͼ": Command3.ToolTipText = "����ǰͼ�񱣴�ΪͼƬ�ļ�"
    Command4.Caption = "¼��": Command4.ToolTipText = "��ʼ¼�񣬱���Ϊ��Ƶ�ļ�"

    H1 = Me.TextHeight("A")
    Command1.Move H1 * 0.5, H1 * 0.5, H1 * 4, H1 * 2
    Command2.Move H1 * 5, H1 * 0.5, H1 * 4, H1 * 2
    Command3.Move H1 * 10, H1 * 0.5, H1 * 4, H1 * 2
    Command4.Move H1 * 15, H1 * 0.5, H1 * 4, H1 * 2
   '�����û�����
    Call ReadSaveSet
    KjEnabled True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call ReadSaveSet(True) '�����û�����
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
   '���������û����õ�ͼƬ����ƵĬ�ϱ���Ŀ¼
    Dim nKey As String, nSub As String
    nKey = "����ͷ���Ƴ���": nSub = "UserOpt"
    If IsSave Then
       SaveSetting nKey, nSub, "AviPath", ctAviPath
       SaveSetting nKey, nSub, "PicPath", ctPicPath
    Else
       ctAviPath = GetSetting(nKey, nSub, "AviPath", "")
       ctPicPath = GetSetting(nKey, nSub, "PicPath", "")
    End If
End Sub
