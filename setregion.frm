VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   Picture         =   "setregion.frx":0000
   ScaleHeight     =   3450
   ScaleWidth      =   6450
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image2 
      Height          =   255
      Index           =   3
      Left            =   5130
      Picture         =   "setregion.frx":31F4
      Top             =   165
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image2 
      Height          =   270
      Index           =   2
      Left            =   4800
      Picture         =   "setregion.frx":3746
      Top             =   180
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Image2 
      Height          =   255
      Index           =   1
      Left            =   4455
      Picture         =   "setregion.frx":3CF0
      Top             =   180
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   270
      Index           =   0
      Left            =   960
      Picture         =   "setregion.frx":4286
      Top             =   150
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   3405
      Left            =   15
      Picture         =   "setregion.frx":47E8
      Top             =   0
      Width           =   6375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
SetRegion
End Sub
Private Sub SetRegion()
Dim hRgn As Long
If hRgn Then DeleteObject hRgn
hRgn = GetBitmapRegion(Me.Picture, vbBlack)
SetWindowRgn Me.hWnd, hRgn, True
End Sub


Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
FormDrag Me
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim num As Integer
For num = 0 To 3
If x >= Image2(num).Left And x <= Image2(num).Left + Image2(num).Width And y >= Image2(num).Top And y <= Image2(num).Top + Image2(num).Height Then
Image2(num).Visible = True
Else
Image2(num).Visible = False
End If
Next num
End Sub



Private Sub Image2_Click(Index As Integer)

Select Case Image2(Index).Index
Case 0
End
Case 1
Me.WindowState = 1
Case 2
Me.WindowState = 0
Case 3
End
End Select



End Sub
