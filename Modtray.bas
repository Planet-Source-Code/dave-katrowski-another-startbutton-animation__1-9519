Attribute VB_Name = "modTray"
Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type

Public NI As NOTIFYICONDATA

Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_TIP = &H4
Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const NIM_MODIFY = &H1

'Button Click Message Constants
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205

Public Const WM_MOUSEMOVE = &H200

Sub Init_Tray(ByRef Frm As Form, Tip As String)
With NI
.cbSize = Len(NI)
.hIcon = Frm.Icon
.hwnd = Frm.hwnd
.szTip = Tip & vbNullChar
.uCallbackMessage = WM_MOUSEMOVE
.uID = 420&
.uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
End With
Shell_NotifyIcon NIM_ADD, NI
End Sub

Sub Unload_Tray()
Shell_NotifyIcon NIM_DELETE, NI
End Sub
