Attribute VB_Name = "Module2"
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Const GW_OWNER = 4
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Boolean
Dim S As String
Dim a As Long
Dim v As Long
S = String(80, 0)
Call GetWindowText(hwnd, S, 80)
S = Left(S, InStr(S, Chr(0)) - 1)
v = GetWindow(hwnd, GW_OWNER)
a = IsWindowVisible(hwnd)
If Len(S) > 0 And a <> 0 And v = 0 Then
Form1.List1.AddItem S
End If
EnumWindowsProc = True
End Function

