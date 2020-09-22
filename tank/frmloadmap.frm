VERSION 5.00
Begin VB.Form frmloadmap 
   BackColor       =   &H00808000&
   ClientHeight    =   2820
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2820
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton load 
      Caption         =   "&Load"
      Height          =   375
      Left            =   2835
      TabIndex        =   1
      Top             =   2445
      Width           =   1815
   End
   Begin VB.ListBox lstMaps 
      Height          =   2235
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmloadmap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub load_Click()
Dim descr, str As String
Dim i, j As Integer

 If lstMaps.ListIndex < 0 Then Exit Sub
 
 'Read file's contents into the relevant variables
 Open App.Path + "\maps\" + arrfile(1 + lstMaps.ListIndex) For Input As #1
  Input #1, descr
  Input #1, grdx
  Input #1, grdy
  ReDim mainarr(1 To 1, 1 To 1)
  ReDim enemytank(1 To 1) As Integer
  enemycount = 0
  ReDim mainarr(1 To grdx, 1 To grdy)
  For j = 1 To grdy
   For i = 1 To grdx
    Input #1, str
    mainarr(i, j) = Val(str)
   Next
  Next
 Close #1
 
 '' Set limits for the enemy tracking array
 ReDim enemytrack(grdx * grdy)
 
 'frmloadmap.Hide
 gameover = False
 Unload Me
 frmGame.drawblocks
 numlives = 4
 frmGame.piclives(1).Visible = True
 frmGame.piclives(2).Visible = True
 frmGame.piclives(3).Visible = True
 frmGame.piclives(4).Visible = True
 frmGame.start.Enabled = True
 frmGame.Show
 
End Sub

Private Sub load_KeyDown(KeyCode As Integer, Shift As Integer)
 frmGame.Text1.SetFocus
End Sub

Private Sub lstMaps_DblClick()
 Call load_Click
End Sub
