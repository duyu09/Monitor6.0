VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "视频捕捉"
   ClientHeight    =   3480
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   4185
   Icon            =   "VBmemcap.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   232
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   279
   StartUpPosition =   2  '屏幕中心
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuExit 
         Caption         =   "退出(E&)"
      End
   End
   Begin VB.Menu mnuControl 
      Caption         =   "控制(&C)"
      Begin VB.Menu mnuStart 
         Caption         =   "开始(&S)"
      End
      Begin VB.Menu mnuDisplay 
         Caption         =   "播放(&D)"
      End
      Begin VB.Menu mnuFormat 
         Caption         =   "设置(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuSource 
         Caption         =   "来源(&o)"
      End
      Begin VB.Menu mnuCompression 
         Caption         =   "压缩(&m)"
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPreview 
         Caption         =   "预览(&P)"
         Checked         =   -1  'True
         Shortcut        =   ^P
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Download by http://www.NewXing.com
'*
'* Author: E. J. Bantz Jr.
'* Copyright: None, use and distribute freely ...
'* E-Mail: ejbantz@usa.net
'* Web: http://www.inlink.com/~ejbantz
'
'* Original C Code Author: Neil Kolban
'* E-Mail: kolban@onramp.net
'* Web: http://rampages.onramp.net/~kolban/video
'*
Option Explicit
Dim lwndC As Long       ' Handle to the Capture Windows
Dim lNFrames As Long  ' Number of frames captured

Sub ResizeCaptureWindow(ByVal lwnd As Long)

    Dim CAPSTATUS As CAPSTATUS
    
    '// Get the capture window attributes .. width and height
    capGetStatus lwnd, VarPtr(CAPSTATUS), Len(CAPSTATUS)
        
    '// Resize the capture window to the capture sizes
    SetWindowPos lwnd, HWND_BOTTOM, 0, 0, _
                       CAPSTATUS.uiImageWidth, _
                       CAPSTATUS.uiImageHeight, _
                       SWP_NOMOVE Or SWP_NOZORDER
         
End Sub


Private Sub Form_Load()

    Dim lpszName As String * 100
    Dim lpszVer As String * 100
    Dim Caps As CAPDRIVERCAPS

    '//Create Capture Window
    capGetDriverDescriptionA 0, lpszName, 100, lpszVer, 100  '// Retrieves driver info
    lwndC = capCreateCaptureWindowA(lpszName, WS_CHILD Or WS_VISIBLE, 0, 0, 160, 120, Me.hwnd, 0)

    '// Connect the capture window to the driver
    capDriverConnect lwndC, 0
    
    '// Get the capabilities of the capture driver
    capDriverGetCaps lwndC, VarPtr(Caps), Len(Caps)
    
    '// If the capture driver does not support a dialog, grey it out
    '// in the menu bar.
    If Caps.fHasDlgVideoSource = 0 Then mnuSource.Enabled = False
    If Caps.fHasDlgVideoFormat = 0 Then mnuFormat.Enabled = False
    If Caps.fHasDlgVideoDisplay = 0 Then mnuDisplay.Enabled = False
    
    '// Set the video stream callback function
    capSetCallbackOnVideoStream lwndC, AddressOf MyVideoStreamCallback
    capSetCallbackOnFrame lwndC, AddressOf MyFrameCallback
    
    '// Set the preview rate in milliseconds
    capPreviewRate lwndC, 66
    
    '// Start previewing the image from the camera
    capPreview lwndC, True
        
    '// Resize the capture window to show the whole image
    ResizeCaptureWindow lwndC

End Sub


Private Sub Form_Unload(Cancel As Integer)

    '// Disable all callbacks
    capSetCallbackOnError lwndC, vbNull
    capSetCallbackOnStatus lwndC, vbNull
    capSetCallbackOnYield lwndC, vbNull
    capSetCallbackOnFrame lwndC, vbNull
    capSetCallbackOnVideoStream lwndC, vbNull
    capSetCallbackOnWaveStream lwndC, vbNull
    capSetCallbackOnCapControl lwndC, vbNull
    

End Sub


Private Sub mnuCompression_Click()
'   /*
'   * Display the Compression dialog when "Compression" is selected from
'   * the menu bar.
'   */
    
    capDlgVideoCompression lwndC

End Sub

Private Sub mnuDisplay_Click()
'   /*
'   * Display the Video Display dialog when "Display" is selected from
'   * the menu bar.
'   */

    capDlgVideoDisplay lwndC
    
End Sub

Private Sub mnuExit_Click()

    Unload Me
    
End Sub

Private Sub mnuFormat_Click()
'  /*
'   * Display the Video Format dialog when "Format" is selected from the
'   * menu bar.
'   */

    capDlgVideoFormat lwndC
    ResizeCaptureWindow lwndC

End Sub

Private Sub mnuPreview_Click()

    mnuPreview.Checked = Not (mnuPreview.Checked)
    capPreview lwndC, mnuPreview.Checked

End Sub


Private Sub mnuSource_Click()
'   /*
'    * Display the Video Source dialog when "Source" is selected from the
'    * menu bar.
'    */
    
    capDlgVideoSource lwndC

End Sub

Private Sub mnuStart_Click()
' /*
'  * If Start is selected from the menu, start Streaming capture.
'  * The streaming capture is terminated when the Escape key is pressed
'  */
    lNFrames = 0
    capCaptureSequenceNoFile lwndC
    
End Sub


