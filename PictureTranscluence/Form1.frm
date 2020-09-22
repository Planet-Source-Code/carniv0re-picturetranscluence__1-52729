VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   ScaleHeight     =   5640
   ScaleWidth      =   6960
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   5655
      Left            =   0
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   5595
      ScaleWidth      =   6915
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   720
         Top             =   2160
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Load Form2
Form2.Show

Picture1.Picture = LoadPicture(Pic1)
Form2.Height = Form1.Height
Form2.Width = Form1.Width
Form2.Top = Form1.Top
Form2.Left = Form1.Left
a = MakeTranslucent(Form1.hWnd, 125)
a = MakeTranslucent(Form2.hWnd, 125)
End Sub

Private Sub Form_Resize()
Form2.Height = Form1.Height
Form2.Width = Form1.Width
Form2.Top = Form1.Top
Form2.Left = Form1.Left
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Form2
End
End Sub

Private Sub Timer1_Timer()
Form2.Height = Form1.Height
Form2.Width = Form1.Width
Form2.Top = Form1.Top
Form2.Left = Form1.Left

If FLAG = True Then
   OPA = OPA - 5
     Else
   OPA = OPA + 5
End If

If OPA >= 255 Then
   FLAG = True
End If

If OPA <= 0 Then
   FLAG = False
End If
a = MakeTranslucent(Form1.hWnd, 255 - OPA)
a = MakeTranslucent(Form2.hWnd, OPA)
End Sub
