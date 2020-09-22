VERSION 5.00
Begin VB.Form frmEditor 
   BackColor       =   &H00808000&
   Caption         =   "Design you own Map:"
   ClientHeight    =   8595
   ClientLeft      =   1155
   ClientTop       =   1230
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton gamescreen 
      Caption         =   "&Game screen"
      Height          =   495
      Left            =   10320
      TabIndex        =   11
      Top             =   1440
      Width           =   1335
   End
   Begin VB.PictureBox yourtankpic 
      Height          =   615
      Left            =   120
      Picture         =   "editor.frx":0000
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   10
      Top             =   7920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox yourtank 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   225
      Picture         =   "editor.frx":069E
      ScaleHeight     =   615
      ScaleWidth      =   495
      TabIndex        =   9
      ToolTipText     =   "Your Tank"
      Top             =   2490
      Width           =   495
   End
   Begin VB.PictureBox tank1pic 
      Height          =   615
      Left            =   120
      Picture         =   "editor.frx":0BF5
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   8
      Top             =   7200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox emptyblock 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   240
      ScaleHeight     =   615
      ScaleWidth      =   495
      TabIndex        =   7
      ToolTipText     =   "Empty Block"
      Top             =   1695
      Width           =   495
   End
   Begin VB.PictureBox tank1 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   240
      Picture         =   "editor.frx":1293
      ScaleHeight     =   615
      ScaleWidth      =   495
      TabIndex        =   6
      ToolTipText     =   "Enemy Tank"
      Top             =   885
      Width           =   495
   End
   Begin VB.PictureBox wall1 
      Height          =   615
      Left            =   240
      Picture         =   "editor.frx":1807
      ScaleHeight     =   555
      ScaleWidth      =   435
      TabIndex        =   5
      ToolTipText     =   "Wall"
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox currentpic 
      Height          =   495
      Left            =   840
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   4
      Top             =   8280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   495
      Left            =   10320
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.PictureBox block1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   1440
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   1
      Top             =   8280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton exit 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   10320
      TabIndex        =   0
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label map 
      BorderStyle     =   1  'Fixed Single
      Height          =   8055
      Left            =   1035
      TabIndex        =   3
      Top             =   105
      Width           =   9135
   End
End
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim counter As Integer
Dim curx, cury As Integer
Dim gridx, gridy As Integer
Dim i, j As Integer
Dim clicked As Boolean
Dim mytankchosen As Boolean
Private Sub block1_Click(Index As Integer)
 
  If block1(Index).Picture = yourtankpic.Picture Then
   mytankchosen = False
   yourtank.Enabled = True
  End If
  If mytankchosen = True And currentpic.Picture = yourtankpic.Picture Then Exit Sub
  block1(Index).ZOrder (0)
  block1(Index).Picture = currentpic.Picture
  'block1(Index).BorderStyle = 1
  If currentpic.Picture = yourtankpic.Picture Then
   mytankchosen = True
   yourtank.Enabled = False
  End If
 
End Sub

Private Sub block1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
 If Shift = 1 Then clicked = True
End Sub

Private Sub block1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
 clicked = False
End Sub

Private Sub block1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

  If clicked = True Then
   If block1(Index).Picture = yourtankpic.Picture Then
    mytankchosen = False
    yourtank.Enabled = True
   End If
   If currentpic.Picture <> yourtankpic.Picture Then
    block1(Index).ZOrder (0)
    block1(Index).Picture = currentpic.Picture
    'block1(Index).BorderStyle = 1
   End If
  End If
  
End Sub

Private Sub cmdSave_Click()
Dim filename, description As String
Dim i As Integer
Dim pathoffile As String
Dim done As Boolean

'The players tank has not been selected
If mytankchosen = False Then
 MsgBox "You must include your tank in the map!", , "Error"
 Exit Sub
End If
filename = InputBox("Enter the file name", "File Name")

done = False
While (done = False)
 description = InputBox("Enter the description for this map", "Description")
 If description = "" Then
  MsgBox "You must enter in a map description", , "Error:"
 Else
  done = True
 End If
Wend

pathoffile = App.Path + "\maps\" + filename + ".map"

Open pathoffile For Append As #1

Write #1, description
Write #1, gridx
Write #1, gridy

For i = 1 To (gridx * gridy)
 If block1(i).Picture = emptyblock.Picture Then Write #1, Int("0")
 If block1(i).Picture = wall1.Picture Then Write #1, Int("1")
 If block1(i).Picture = tank1pic.Picture Then Write #1, Int("2")
 If block1(i).Picture = yourtankpic.Picture Then Write #1, Int("3")
Next
Close #1

End Sub

Private Sub emptyblock_Click()

 yourtank.BorderStyle = 0
 wall1.BorderStyle = 0
 tank1.BorderStyle = 0
 emptyblock.BorderStyle = 1
 currentpic.Picture = emptyblock.Picture
 
End Sub

Private Sub emptyblock_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 frmEditor.Caption = "Design your Own Map - Empty Block"
End Sub

Private Sub exit_Click()
 End
End Sub

Private Sub Form_Activate()

 clicked = False
 mytankchosen = False
 counter = 0
 currentpic.Picture = wall1.Picture
 emptyblock.BackColor = map.BackColor
 'number of blocks:
 gridx = 10
 gridy = 10
 
 'width and height of each block is worked out
 block1(0).Width = map.Width / gridx
 block1(0).Height = map.Height / gridy
 
 'draws blocks
 For i = 0 To gridy - 1
  For j = 0 To gridx - 1
   counter = counter + 1
   load block1(counter)
   block1(counter).Left = j * block1(counter).Width + map.Left
   block1(counter).Top = i * block1(counter).Height + map.Top
   block1(counter).Visible = True
  Next
 Next
 
End Sub


Private Sub Form_Unload(Cancel As Integer)
 For i = 1 To counter
  Unload block1(i)
 Next
End Sub

Private Sub gamescreen_Click()

 frmGame.Show
 frmEditor.Hide
 
End Sub

Private Sub tank1_Click()

 yourtank.BorderStyle = 0
 emptyblock.BorderStyle = 0
 wall1.BorderStyle = 0
 tank1.BorderStyle = 1
 currentpic.Picture = tank1pic.Picture
 
End Sub

Private Sub tank1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 frmEditor.Caption = "Design your Own Map - Enemy Tank"
End Sub

Private Sub wall1_Click()

 yourtank.BorderStyle = 0
 emptyblock.BorderStyle = 0
 tank1.BorderStyle = 0
 wall1.BorderStyle = 1
 currentpic.Picture = wall1.Picture
 
End Sub

Private Sub wall1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 frmEditor.Caption = "Design your Own Map - Wall"
End Sub

Private Sub yourtank_Click()

 emptyblock.BorderStyle = 0
 tank1.BorderStyle = 0
 wall1.BorderStyle = 0
 yourtank.BorderStyle = 1
 currentpic.Picture = yourtankpic.Picture

End Sub

Private Sub yourtank_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 frmEditor.Caption = "Design your Own Map - Your Tank"
End Sub
