Attribute VB_Name = "modSysTray"
Option Explicit

' Declare the API function call.
Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

' Declare a user-defined variable to pass to the Shell_NotifyIcon function.
Public Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 128
    dwState As Long
    dwStateMask As Long
    szInfo As String * 256
    uTimeout As Long
    szInfoTitle As String * 64
    dwInfoFlags As Long
End Type

' Define a variable to the type.
Public NID As NOTIFYICONDATA

' Messages sent to the Notify Shell command.
Public Const NIS_HIDDEN As Long = &H1
Public Const NIS_SHAREDICON As Long = &H2

' Messages sent to the Notify Shell command.
Public Const NIM_ADD As Long = &H0
Public Const NIM_MODIFY As Long = &H1
Public Const NIM_DELETE As Long = &H2

' Message types sent to the Notify Shell command.
Public Const NIF_MESSAGE As Long = &H1
Public Const NIF_ICON As Long = &H2
Public Const NIF_TIP As Long = &H4
Public Const NIF_STATE As Long = &H8
Public Const NIF_INFO As Long = &H10

' Left-click constants.
Public Const WM_MOUSEMOVE As Long = &H200
Public Const WM_LBUTTONDOWN As Long = &H201
Public Const WM_LBUTTONUP As Long = &H202
Public Const WM_LBUTTONDBLCLK As Long = &H203

' Right-click constants.
Public Const WM_RBUTTONDBLCLK As Long = &H206
Public Const WM_RBUTTONDOWN As Long = &H204
Public Const WM_RBUTTONUP As Long = &H205
