Attribute VB_Name = "notify"
Option Explicit
Private Declare PtrSafe Function Shell_NotifyIconA Lib "shell32.dll" (ByVal dwMessage As Long, ByRef nfIconData As NOTIFYICONDATA) As LongPtr
Public nfIconData As NOTIFYICONDATA
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As LongPtr
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As LongPtr
    szTip As String * 128
    dwState As Long
    dwStateMask As Long
    szInfo As String * 256
    uTimeout As Long
    szInfoTitle As String * 64
    dwInfoFlags As Long
End Type
Public Function toast(Optional ByVal title As String, Optional ByVal info As String, Optional ByVal flag As Long)
With nfIconData
    .dwInfoFlags = flag
    .uFlags = &H10
    .szInfoTitle = title
    .szInfo = info
    .cbSize = &H1F8
End With
Shell_NotifyIconA &H0, nfIconData
Shell_NotifyIconA &H1, nfIconData
End Function
'Flags for the balloon message..
'None = 0
'Information = 1
'Exclamation = 2
'Critical = 3
