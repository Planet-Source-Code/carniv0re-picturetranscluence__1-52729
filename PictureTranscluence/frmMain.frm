VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PictureTranscluence"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7170
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7170
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   3600
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1080
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdChange1 
      Caption         =   "Load Picture"
      BeginProperty Font 
         Name            =   "Narkisim"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   3000
      Width           =   3495
   End
   Begin VB.CommandButton cmdTwo 
      Caption         =   "Load Picture"
      BeginProperty Font 
         Name            =   "Narkisim"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      Top             =   3000
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   435
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   3480
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   435
      Left            =   3600
      TabIndex        =   2
      Top             =   0
      Width           =   3480
   End
   Begin VB.Image Picture1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2415
      Left            =   0
      Stretch         =   -1  'True
      Top             =   480
      Width           =   3495
   End
   Begin VB.Image Picture2 
      BorderStyle     =   1  'Fixed Single
      Height          =   2415
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   480
      Width           =   3495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdChange1_Click()
On Error GoTo r:
CommonDialog1.ShowOpen
Picture1.Picture = LoadPicture(CommonDialog1.FileName)
Label1.Caption = CommonDialog1.FileTitle
Pic1 = CommonDialog1.FileName
Exit Sub
r:
MsgBox ("Invalid Picture!!!")
End Sub

Private Sub cmdStart_Click()
frmMain.Hide
Load Form1
Form1.Show
End Sub

Private Sub cmdTwo_Click()
On Error GoTo r:
CommonDialog1.ShowOpen
Picture2.Picture = LoadPicture(CommonDialog1.FileName)
Label1.Caption = CommonDialog1.FileTitle
Pic2 = CommonDialog1.FileName
Exit Sub
r:
MsgBox ("Invalid Picture!!!")
End Sub
