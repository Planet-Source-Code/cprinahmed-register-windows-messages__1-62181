VERSION 5.00
Begin VB.Form Receiver 
   Caption         =   "Receiver"
   ClientHeight    =   1275
   ClientLeft      =   3450
   ClientTop       =   2505
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   ScaleHeight     =   1275
   ScaleWidth      =   5055
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "Receiver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40

Private Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Private Sub Form_Activate()
    
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE

End Sub

Private Sub Form_Load()
    
    SubClass Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
     
    UnSubClass Me

End Sub



