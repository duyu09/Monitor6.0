VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Duyu - ���������"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6270
   Icon            =   "3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   6270
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command4 
      BackColor       =   &H80000003&
      Caption         =   "�����½���"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ListBox List1 
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5190
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   240
      Width           =   4695
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000003&
      Caption         =   "�ر�"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "����ѡ�еĽ���"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4920
      MaskColor       =   &H000000FF&
      TabIndex        =   1
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000003&
      Caption         =   "ˢ��"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'----------------------------------------- API���� -------------------------------------------------------
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function Module32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As MODULEENTRY32) As Long
Private Declare Function Module32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As MODULEENTRY32) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Thread32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As THREADENTRY32) As Long
Private Declare Function Thread32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As THREADENTRY32) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
'---------------------------------------- API�������� ------------------------------------------------------
Private Const TH32CS_SNAPHEAPLIST = &H1
Private Const TH32CS_SNAPPROCESS = &H2
Private Const TH32CS_SNAPTHREAD = &H4
Private Const TH32CS_SNAPMODULE = &H8
Private Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Private Const TH32CS_INHERIT = &H80000000
Private Const PROCESS_TERMINATE = &H1&
'---------- API�������� -----------
Private Type PROCESSENTRY32 '����
dwSize As Long
cntusage As Long
th32ProcessID As Long
th32DefaultHeapID As Long
th32ModuleID As Long
cntThreads As Long
th32ParentProcessID As Long
pcPriClassBase As Long
dwFlags As Long
szExeFile As String * 1024
End Type
Private Type MODULEENTRY32 'ģ��
dwSize As Long
th32ModuleID As Long
th32ProcessID As Long
GlblcntUsage As Long
ProccntUsage As Long
modBaseAddr As Byte
modBaseSize As Long
hModule As Long
szModule As String * 256
szExePath As String * 1024
End Type
Private Type THREADENTRY32 '�߳�
dwSize As Long
cntusage As Long
th32threadID As Long
th32OwnerProcessID As Long
tpBasePri As Long
tpDeltaPri As Long
dwFlags As Long
End Type
Dim ProcessID() As Long '��list1�еĽ�˳��洢���н���ID


Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const HWND_TOPMOST& = -1
' �����������б�������λ���κ�������ڵ�ǰ��
Private Const SWP_NOSIZE& = &H1
' ���ִ��ڴ�С
Private Const SWP_NOMOVE& = &H2
' ���ִ���λ��

Private Sub Command1_Click()
Dim Process As PROCESSENTRY32
Dim ProcSnap As Long
Dim cntProcess As Long
cntProcess = 0
List1.Clear
ProcSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
If ProcSnap Then
Process.dwSize = 1060 ' ͨ���÷�
Process32First ProcSnap, Process
Do Until Process32Next(ProcSnap, Process) < 1 ' �������н���ֱ������ֵΪFalse
List1.AddItem Trim(Process.szExeFile)
cntProcess = cntProcess + 1
Loop
End If
ReDim ProcessID(cntProcess) As Long
Dim i As Long
i = 0
Process32First ProcSnap, Process
Do Until Process32Next(ProcSnap, Process) < 1 ' �������н���ֱ������ֵΪFalse
ProcessID(i) = Process.th32ProcessID
i = i + 1
Loop
CloseHandle (ProcSnap)
End Sub
Private Sub Command2_Click()
Dim c As Integer, ssre As Integer, g As String
If List1.ListIndex < 0 Then
MsgBox "��ѡ�����!", vbOKOnly + vbInformation, "��ʾ"
Else
g = List1.Text
Dim hProcess As Long
hProcess = OpenProcess(PROCESS_TERMINATE, False, ProcessID(List1.ListIndex))
If hProcess Then TerminateProcess hProcess, 0
c = List1.ListCount
While List1.ListCount = c
Command1_Click
ssre = MsgBox("���̽���ʧ��.�뵥�������ԡ� ǿ�ƽ���.", vbRetryCancel)
If ssre = 4 Then
Shell ("taskkill /f /im " & g)
Command1_Click
End If
Command1_Click
Wend
Command1_Click
End If
Command1_Click
End Sub
Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
Dim stk As String, ss As Integer
stk = InputBox("�������������������������������.")
On Error Resume Next
ss = Shell(Chr(34) & stk & Chr(34), vbNormalFocus)
If Err.Number > 0 Then
ss = MsgBox(Err.Description, 48)
End If
End Sub

Private Sub Form_Load()
Command1_Click
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
' ��������Ϊ������ǰ
End Sub

Private Sub Timer1_Timer()

End Sub
