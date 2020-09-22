Attribute VB_Name = "Module1"
Option Explicit

Private WM_MyOwnMessage As Long

Declare Function GetProp Lib "User32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Declare Function RemoveProp Lib "User32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SetProp Lib "User32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function CallWindowProc Lib "User32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


Public Const GWL_WNDPROC = (-4)

Dim PrevProc As Long

Public Sub SubClass(F As Form)
    
    PrevProc = SetWindowLong(F.hWnd, GWL_WNDPROC, AddressOf WindowProc)

End Sub
Public Sub UnSubClass(F As Form)
    
    SetWindowLong F.hWnd, GWL_WNDPROC, PrevProc

End Sub
Public Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  
Static c As Long

Dim SenderHwnd As Long
  
  SenderHwnd = FindWindow(vbNullString, "Sender")
  
  WM_MyOwnMessage = GetProp(SenderHwnd, "messages")
  
  If uMsg = WM_MyOwnMessage Then
     c = c + 1
     Debug.Print "WM_MyOwnMessage Received from hwnd " & SenderHwnd
     Receiver.Text1 = c & ".WM_MyOwnMessage Received from Sender Hwnd " & SenderHwnd
     
     If wParam = 4 Then
         
        MsgBox "hi sender"
         
     ElseIf wParam = 5 Then
     
        MsgBox "fine"
        
     End If
     
  End If
  
   WindowProc = CallWindowProc(PrevProc, hWnd, uMsg, wParam, lParam)
  

End Function

