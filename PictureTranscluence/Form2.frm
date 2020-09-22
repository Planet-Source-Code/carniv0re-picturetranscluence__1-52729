VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7425
   LinkTopic       =   "Form2"
   ScaleHeight     =   5850
   ScaleWidth      =   7425
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   5895
      Left            =   0
      ScaleHeight     =   5835
      ScaleWidth      =   7395
      TabIndex        =   0
      Top             =   0
      Width           =   7455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Picture1.Picture = LoadPicture(Pic2)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Form1
End
End Sub

Private Sub Picture1_Resize()
Form1.Height = Form2.Height
Form1.Width = Form2.Width
Form1.Top = Form2.Top
Form1.Left = Form2.Left
End Sub
