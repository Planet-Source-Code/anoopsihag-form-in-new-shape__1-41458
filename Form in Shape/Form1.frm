VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Shape"
   ClientHeight    =   3300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4125
   DrawMode        =   10  'Mask Pen
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   4125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   2265
      Left            =   0
      Picture         =   "Form1.frx":0000
      ToolTipText     =   "Click for Exit"
      Top             =   0
      Width           =   2745
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long


Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long


Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long

Private Const RGN_COPY = 5
Private Const SWP_NOMOVE = &H2

Private Const HWND_TOPMOST = -1
Private dx As Integer, dy As Integer, dwn As Integer



Private Sub Form_Load()
Dim rgn As Long
Dim rgn2 As Long
SetWindowPos Form1.hwnd, HWND_TOPMOST, 100, 100, 800, 800, 0
rgn = CreateEllipticRgn(81, 51, 180, 150)
rgn2 = CreateRectRgn(90, 50, 180, 150)
 CombineRgn rgn, rgn, rgn2, RGN_COPY
SetWindowRgn Form1.hwnd, rgn, True

End Sub

Private Sub Image1_DblClick()
MsgBox "             Thanx         " + vbCrLf + "anoopsihag@yahoo.co.in"
End
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
dx = x
dy = y
dwn = True
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If dwn Then
    Move Left + (x - dx), Top + (y - dy)
    u% = DoEvents 'make sure it cleans up
End If
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
dwn = False
End Sub
