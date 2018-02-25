VERSION 5.00
Begin VB.Form frmSystray 
   Caption         =   "Form1"
   ClientHeight    =   735
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   1560
   Icon            =   "frmSystemTray.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   735
   ScaleWidth      =   1560
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox picFlash1 
      Height          =   660
      Left            =   0
      Picture         =   "frmSystemTray.frx":1CFA
      ScaleHeight     =   600
      ScaleWidth      =   870
      TabIndex        =   0
      Top             =   0
      Width           =   930
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuConfigure 
         Caption         =   "&Configure"
      End
      Begin VB.Menu mnuShowTips 
         Caption         =   "&Show Tips"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmSystray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Private Const WM_MOUSEMOVE = &H200
Private Const NIF_ICON = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_TIP = &H4
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const MAX_TOOLTIP As Integer = 64
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * MAX_TOOLTIP
End Type
Private nfIconData As NOTIFYICONDATA

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDOWN = &H204

Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Private Const SW_SHOWMINIMIZED = 2
Private Const SW_NORMAL = 1

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    App.TaskVisible = False
    SetIcon
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lX, WindowHandle As Long
   lX = ScaleX(X, Me.ScaleMode, vbPixels)
    Select Case lX
        Case WM_LBUTTONDBLCLK
            frmMain.Show
'            WindowHandle = FindWindow(vbNullString, "SystemTray In Java") '
'            If (WindowHandle <> 0) Then
'                BringWindowToTop WindowHandle
'                ShowWindow WindowHandle, SW_SHOWMINIMIZED
'                ShowWindow WindowHandle, SW_NORMAL
'            End If
        Case WM_RBUTTONDOWN
            Call Me.PopupMenu(mnuFile)

    End Select
End Sub
Public Sub SetIcon()
    With nfIconData
        .hwnd = Me.hwnd
        .uID = Me.Icon
        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = picFlash1.Picture
        .szTip = "System Tray Demo in Java" & Chr$(0)
        .cbSize = Len(nfIconData)
    End With
    Shell_NotifyIcon NIM_ADD, nfIconData
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Shell_NotifyIcon NIM_DELETE, nfIconData
End Sub



Private Sub mnuExit_Click()
    Unload Me
End Sub

