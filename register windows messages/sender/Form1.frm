VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Sender"
   ClientHeight    =   2070
   ClientLeft      =   4395
   ClientTop       =   1470
   ClientWidth     =   3105
   LinkTopic       =   "Form1"
   ScaleHeight     =   2070
   ScaleWidth      =   3105
   Begin VB.CommandButton Command1 
      Caption         =   "Say How are you?"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   2655
   End
   Begin VB.CommandButton cmdCloseRece 
      Caption         =   "Say hi"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   2655
   End
   Begin VB.CommandButton cmdSendMessage 
      Caption         =   "Send Message To  Receiver"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WM_MyOwnMessage As Long

Private Declare Function GetProp Lib "User32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function RegisterWindowMessage Lib "User32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Private Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetProp Lib "User32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function PostMessage Lib "User32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wparam As Long, lparam As Any) As Long
Private Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40

Private Sub cmdCloseRece_Click()

SendMessageToRece 4, 0

End Sub

Private Sub Command1_Click()

SendMessageToRece 5, 0

End Sub

Private Sub Form_Activate()
   
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE

End Sub

Private Sub cmdSendMessage_Click()

SendMessageToRece 0, 0

End Sub

Private Sub Form_Load()

Dim Ok As Boolean

Ok = RegisterMessage

If Not Ok Then
   
   MsgBox "Error Registering Your message"

End If

End Sub

Private Function RegisterMessage() As Boolean
    
    If WM_MyOwnMessage = 0 Then
      
      WM_MyOwnMessage = RegisterWindowMessage("MyOwnMessage")
      
      If WM_MyOwnMessage <> 0 Then
         
         RegisterMessage = True
         
         SetProp hWnd, "messages", WM_MyOwnMessage
      
      End If
    
    End If
    
End Function



Public Sub SendMessageToRece(Optional wparam As Long, Optional lparam As Long)

Dim Res As Long

Dim ReceiverHwnd As Long

ReceiverHwnd = FindWindow(vbNullString, "Receiver")

Res = PostMessage(ReceiverHwnd, WM_MyOwnMessage, wparam, lparam)

If Res = 0 Then
   
   MsgBox "Error sending messages"

End If

End Sub
