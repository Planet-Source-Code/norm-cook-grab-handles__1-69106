VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Test Grabhandles"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin prjGrabHandles.GH GH1 
      Height          =   315
      Left            =   2160
      TabIndex        =   2
      Top             =   1440
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Toggle Resize Command1"
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   2640
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Resize Me"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
 If GH1.Sizing Then
  GH1.Detach
 End If
End Sub

Private Sub Command2_Click()
 If GH1.Sizing = False Then
  GH1.Attach Command1
 Else
  GH1.Detach
 End If
End Sub

Private Sub Form_Load()
' GH1.Attach Command1
End Sub
