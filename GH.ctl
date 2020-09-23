VERSION 5.00
Begin VB.UserControl GH 
   ClientHeight    =   345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   315
   ScaleHeight     =   345
   ScaleWidth      =   315
   ToolboxBitmap   =   "GH.ctx":0000
   Begin VB.PictureBox GHInv 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   105
      Left            =   600
      Picture         =   "GH.ctx":0312
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   9
      Top             =   480
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.PictureBox GHObv 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   105
      Left            =   240
      Picture         =   "GH.ctx":03FC
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.PictureBox GH 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   105
      Index           =   7
      Left            =   1800
      Picture         =   "GH.ctx":04E6
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.PictureBox GH 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   105
      Index           =   6
      Left            =   1560
      Picture         =   "GH.ctx":05D0
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.PictureBox GH 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   105
      Index           =   5
      Left            =   1320
      Picture         =   "GH.ctx":06BA
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.PictureBox GH 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   105
      Index           =   4
      Left            =   1080
      Picture         =   "GH.ctx":07A4
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.PictureBox GH 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   105
      Index           =   3
      Left            =   840
      Picture         =   "GH.ctx":088E
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.PictureBox GH 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   105
      Index           =   2
      Left            =   600
      Picture         =   "GH.ctx":0978
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.PictureBox GH 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   105
      Index           =   1
      Left            =   360
      Picture         =   "GH.ctx":0A62
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.PictureBox GH 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   105
      Index           =   0
      Left            =   120
      Picture         =   "GH.ctx":0B4C
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   105
   End
End
Attribute VB_Name = "GH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const GHDIM As Long = 105 'grab handle hgt/wid
Private Const GHHALF As Long = 52
Private Const UL As Long = 0 'upper left
Private Const UR As Long = 1 'upper right
Private Const LL As Long = 2 'lower left
Private Const LR As Long = 3 'lower right
Private Const ML As Long = 4 'middle left
Private Const MT As Long = 5 'middle top
Private Const MR As Long = 6 'middle right
Private Const MB As Long = 7 'middle bottom
Private Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private GX(7) As Long, GY(7) As Long 'for dragging
Private TWx As Long, TWy As Long 'twips per pixel
Private TheCtl As Control 'the sizeable control
Private PHWnd As Long 'parent hwnd

Private Sub UserControl_Resize()
 Static Busy As Boolean
 TWx = Screen.TwipsPerPixelX
 TWy = Screen.TwipsPerPixelY
 'keep the control small & unobtrusive
 If Not Busy Then
  Busy = True 'to avoid recursion
  UserControl.Width = 315
  UserControl.Height = 315
  Busy = False
 End If
End Sub
Public Property Get Sizing() As Boolean
 Sizing = GH(UL).Visible
End Property
Public Sub Attach(Ctl As Control)
 Dim i As Long
 PHWnd = UserControl.ContainerHwnd 'form hwnd
 For i = 0 To 7 'put the grabhandle pics on the form
  SetParent GH(i).hWnd, PHWnd
 Next
 Set TheCtl = Ctl 'makes the control visible to the uc
 MoveGH 'position the pics
 GHVis True ' & show
 UserControl.Extender.Left = -315
 UserControl.Extender.Top = -315
 'UserControl.Extender.ZOrder vbSendToBack
End Sub
Public Sub Detach() 'basically the opposite
 Dim i As Long
 If PHWnd Then
  For i = 0 To 7
   SetParent PHWnd, GH(i).hWnd
  Next
  Set TheCtl = Nothing
  GHVis False
 End If
End Sub
Private Sub MoveGH()
 'position the 8 pics around the control
 With TheCtl
  GH(UL).Move .Left - GHDIM, .Top - GHDIM
  GH(UR).Move .Left + .Width, .Top - GHDIM
  GH(LL).Move .Left - GHDIM, .Top + .Height
  GH(LR).Move .Left + .Width, .Top + .Height

  GH(ML).Move .Left - GHDIM, .Top + (.Height \ 2) - GHHALF
  GH(MT).Move .Left + (.Width \ 2) - GHHALF, .Top - GHDIM
  GH(MR).Move .Left + .Width, .Top + (.Height \ 2) - GHHALF
  GH(MB).Move .Left + (.Width \ 2) - GHHALF, .Top + .Height
 End With
End Sub
Private Sub GH_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 GH(Index).Picture = GHInv.Picture 'show inverse grabhandle pic
 GX(Index) = X \ TWx 'and save mousedown x/y
 GY(Index) = Y \ TWy
End Sub

Private Sub GH_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button Then
  Select Case Index
   Case 4, 6
    'only needs to move left/right
    GH(Index).Move GH(Index).Left + X \ TWx - GX(Index), GH(Index).Top
   Case 5, 7
    'only needs to move up/down
    GH(Index).Move GH(Index).Left, GH(Index).Top + Y \ TWy - GY(Index)
   Case Else
    'can move in any direction
    GH(Index).Move GH(Index).Left + X \ TWx - GX(Index), GH(Index).Top + Y \ TWy - GY(Index)
  End Select
  SizeCtl
 End If
End Sub
Private Sub GH_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 'restore the normal grabhandle pic
 GH(Index).Picture = GHObv.Picture
End Sub
Private Sub GHVis(ByVal Vis As Boolean)
 Dim i As Long
 For i = 0 To 7
  GH(i).Visible = Vis
 Next
End Sub
Private Function NewLeft() As Long
 With TheCtl
  If GH(UL).Left + GHDIM <> .Left Then
   NewLeft = GH(UL).Left + GHDIM
  ElseIf GH(LL).Left + GHDIM <> .Left Then
   NewLeft = GH(LL).Left + GHDIM
  ElseIf GH(ML).Left + GHDIM <> .Left Then
   NewLeft = GH(ML).Left + GHDIM
  End If
 End With
End Function
Private Function NewTop() As Long
 With TheCtl
  If GH(UL).Top + GHDIM <> .Top Then
   NewTop = GH(UL).Top + GHDIM
  ElseIf GH(MT).Top + GHDIM <> .Top Then
   NewTop = GH(MT).Top + GHDIM
  ElseIf GH(UR).Top + GHDIM <> .Top Then
   NewTop = GH(UR).Top + GHDIM
  End If
 End With
End Function
Private Function NewRight() As Long
 With TheCtl
  If GH(UR).Left <> .Left + .Width Then
   NewRight = GH(UR).Left - (.Left + .Width)
  ElseIf GH(MR).Left <> .Left + .Width Then
   NewRight = GH(MR).Left - (.Left + .Width)
  ElseIf GH(LR).Left <> .Left + .Width Then
   NewRight = GH(LR).Left - (.Left + .Width)
  End If
 End With
End Function
Private Function NewBottom() As Long
 With TheCtl
  If GH(LR).Top <> .Top + .Height Then
   NewBottom = GH(LR).Top - (.Top + .Height)
  ElseIf GH(MB).Top <> .Top + .Height Then
   NewBottom = GH(MB).Top - (.Top + .Height)
  ElseIf GH(LL).Top <> .Top + .Height Then
   NewBottom = GH(LL).Top - (.Top + .Height)
  End If
 End With
End Function
Private Sub SizeCtl()
 Dim nl As Long, nt As Long
 Dim nr As Long, nb As Long
 'a lot of trial & error code here
 ' could probably be improved
 With TheCtl
  nl = NewLeft
  If nl Then
   .Left = nl
'   .Width = .Width + .Left - nl
  End If
  nt = NewTop
  If nt Then
 '  .Height = .Height + .Top - nt
   .Top = nt
  End If
  nr = NewRight
  If nr Then
   .Width = .Width + nr
  End If
  nb = NewBottom
  If nb Then
   .Height = .Height + nb
  End If
 End With
 MoveGH 'move the other grabhandles to proper pos
End Sub

