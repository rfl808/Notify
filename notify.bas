Attribute VB_Name = "notify"
Option Explicit

Private Declare PtrSafe Function Shell_NotifyIconW Lib "shell32.dll" (ByVal dwMessage As Long, ByRef nfIconData As NOTIFYICONDATAW) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Private Type NOTIFYICONDATAW
  cbSize As Long
#If Win64 Then
  padding1 As Long
#End If
  hwnd As LongPtr
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
#If Win64 Then
  padding2 As Long
#End If
  hIcon As LongPtr
  szTip(1 To 128 * 2) As Byte
  dwState As Long
  dwStateMask As Long
  szInfo(1 To 256 * 2) As Byte
  uTimeout As Long
  szInfoTitle(1 To 64 * 2) As Byte
  dwInfoFlags As Long
End Type

Private Const NIM_ADD As Long = &H0&
Private Const NIM_MODIFY As Long = &H1&
Private Const NIF_INFO As Long = &H10&

Private Function Min(ByVal a As Long, ByVal b As Long) As Long
  If a < b Then Min = a Else Min = b
End Function

Public Sub Toast(Optional ByVal title As String, Optional ByVal info As String, Optional ByVal flag As Long)
  Dim nfIconData As NOTIFYICONDATAW
  
  info = info & " "
  title = title & " "
  With nfIconData
    .cbSize = Len(nfIconData)
    
    .uFlags = NIF_INFO
    .dwInfoFlags = flag
    
    If Len(title) > 0 Then
      CopyMemory ByVal VarPtr(.szInfoTitle(LBound(.szInfoTitle))), ByVal StrPtr(title), Min(Len(title) * 2, UBound(.szInfoTitle) - LBound(.szInfoTitle) + 1 - 2)
    End If
    
    If Len(info) > 0 Then
      CopyMemory ByVal VarPtr(.szInfo(LBound(.szInfo))), ByVal StrPtr(info), Min(Len(info) * 2, UBound(.szInfo) - LBound(.szInfo) + 1 - 2)
    End If
  End With
  
  Shell_NotifyIconW NIM_ADD, nfIconData
  Shell_NotifyIconW NIM_MODIFY, nfIconData
End Sub


'Flags for the balloon message..
'None = 0
'Information = 1
'Exclamation = 2
'Critical = 3



