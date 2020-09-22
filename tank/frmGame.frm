VERSION 5.00
Begin VB.Form frmGame 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tank Game"
   ClientHeight    =   8070
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   11805
   Begin VB.PictureBox aimethod 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   0
      Left            =   9600
      Picture         =   "frmGame.frx":0000
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   31
      Top             =   5640
      Width           =   405
   End
   Begin VB.PictureBox aimethod 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   1
      Left            =   9600
      Picture         =   "frmGame.frx":038B
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   30
      Top             =   6240
      Width           =   405
   End
   Begin VB.PictureBox check 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   10320
      Picture         =   "frmGame.frx":06FF
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   29
      Top             =   8040
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.PictureBox uncheck 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   9840
      Picture         =   "frmGame.frx":0A8A
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   28
      Top             =   8040
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.CheckBox avoid 
      BackColor       =   &H00808000&
      Caption         =   "Avoid Missiles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   9600
      TabIndex        =   27
      Top             =   6840
      Width           =   1935
   End
   Begin VB.Timer blowupplayer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5880
      Top             =   480
   End
   Begin VB.PictureBox piclives 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   4
      Left            =   10680
      Picture         =   "frmGame.frx":0DFE
      ScaleHeight     =   570
      ScaleWidth      =   420
      TabIndex        =   26
      Top             =   4320
      Width           =   420
   End
   Begin VB.PictureBox piclives 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   3
      Left            =   9960
      Picture         =   "frmGame.frx":1355
      ScaleHeight     =   570
      ScaleWidth      =   420
      TabIndex        =   25
      Top             =   4320
      Width           =   420
   End
   Begin VB.PictureBox piclives 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   2
      Left            =   10680
      Picture         =   "frmGame.frx":18AC
      ScaleHeight     =   570
      ScaleWidth      =   420
      TabIndex        =   24
      Top             =   3600
      Width           =   420
   End
   Begin VB.PictureBox piclives 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   1
      Left            =   9960
      Picture         =   "frmGame.frx":1E03
      ScaleHeight     =   570
      ScaleWidth      =   420
      TabIndex        =   23
      Top             =   3600
      Width           =   420
   End
   Begin VB.Timer timmissileenemy 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   50
      Left            =   5400
      Top             =   480
   End
   Begin VB.PictureBox missileenemy 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   360
      Index           =   0
      Left            =   0
      ScaleHeight     =   360
      ScaleWidth      =   180
      TabIndex        =   21
      Top             =   120
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox enemydown 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   2640
      Picture         =   "frmGame.frx":235A
      ScaleHeight     =   615
      ScaleWidth      =   450
      TabIndex        =   20
      Top             =   975
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox enemyleft 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   1830
      Picture         =   "frmGame.frx":28C0
      ScaleHeight     =   450
      ScaleWidth      =   615
      TabIndex        =   19
      Top             =   1020
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox enemyright 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   1110
      Picture         =   "frmGame.frx":2E06
      ScaleHeight     =   450
      ScaleWidth      =   615
      TabIndex        =   18
      Top             =   1005
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox enemyup 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   495
      Picture         =   "frmGame.frx":334D
      ScaleHeight     =   615
      ScaleWidth      =   450
      TabIndex        =   17
      Top             =   840
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox missileplayer 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   90
      ScaleHeight     =   360
      ScaleWidth      =   180
      TabIndex        =   16
      Top             =   2490
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox tankup 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   420
      Picture         =   "frmGame.frx":38C1
      ScaleHeight     =   615
      ScaleWidth      =   450
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Timer timEnemy 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   50
      Left            =   6570
      Top             =   0
   End
   Begin VB.Timer timBomb 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   100
      Left            =   6105
      Top             =   0
   End
   Begin VB.Timer timplayer 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   5640
      Top             =   0
   End
   Begin VB.CommandButton endgame 
      Caption         =   "&End game"
      Enabled         =   0   'False
      Height          =   495
      Left            =   9690
      TabIndex        =   13
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton start 
      Caption         =   "&Start game"
      Enabled         =   0   'False
      Height          =   495
      Left            =   9690
      TabIndex        =   12
      Top             =   855
      Width           =   1575
   End
   Begin VB.Timer update 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   5205
      Top             =   15
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   5040
      TabIndex        =   0
      Top             =   8520
      Width           =   1290
   End
   Begin VB.PictureBox tankleft 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   3930
      Picture         =   "frmGame.frx":3E4F
      ScaleHeight     =   450
      ScaleWidth      =   615
      TabIndex        =   11
      Top             =   180
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox tankdown 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   4845
      Picture         =   "frmGame.frx":43A1
      ScaleHeight     =   615
      ScaleWidth      =   450
      TabIndex        =   10
      Top             =   195
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox tankright 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   2880
      Picture         =   "frmGame.frx":491F
      ScaleHeight     =   450
      ScaleWidth      =   615
      TabIndex        =   9
      Top             =   165
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox emptyblock 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   2670
      ScaleHeight     =   615
      ScaleWidth      =   495
      TabIndex        =   8
      Top             =   8220
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox wall1 
      Height          =   615
      Left            =   2115
      Picture         =   "frmGame.frx":4E7B
      ScaleHeight     =   555
      ScaleWidth      =   435
      TabIndex        =   7
      Top             =   8235
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox currentpic 
      Height          =   495
      Left            =   1560
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   6
      Top             =   8355
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox tank1pic 
      AutoSize        =   -1  'True
      Height          =   675
      Left            =   855
      Picture         =   "frmGame.frx":BC46
      ScaleHeight     =   615
      ScaleWidth      =   450
      TabIndex        =   5
      Top             =   8280
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox yourtankpic 
      Height          =   615
      Left            =   180
      Picture         =   "frmGame.frx":C1BA
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   4
      Top             =   8295
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton exit 
      Caption         =   "E&xit the game"
      Height          =   495
      Left            =   9690
      TabIndex        =   3
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton cmdload 
      Caption         =   "&Load new game"
      Height          =   495
      Left            =   9690
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox block1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   3255
      ScaleHeight     =   495
      ScaleWidth      =   1695
      TabIndex        =   1
      Top             =   8355
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Shortest Axis Method"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   10080
      TabIndex        =   33
      Top             =   6240
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Random Method"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   10080
      TabIndex        =   32
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lives Remaining:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   9840
      TabIndex        =   22
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00C0C000&
      FillStyle       =   0  'Solid
      Height          =   2055
      Left            =   9600
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Image missright 
      Height          =   180
      Left            =   2205
      Picture         =   "frmGame.frx":C858
      Top             =   15
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image missleft 
      Height          =   180
      Left            =   2175
      Picture         =   "frmGame.frx":C922
      Top             =   195
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image missup 
      Height          =   480
      Left            =   1965
      Picture         =   "frmGame.frx":C9EB
      Top             =   30
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image missdown 
      Height          =   480
      Left            =   2640
      Picture         =   "frmGame.frx":CAC9
      Top             =   120
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label map 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   8055
      Left            =   60
      TabIndex        =   15
      Top             =   60
      Width           =   9135
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i, j As Integer
Dim gamestarted As Boolean
Dim dirplayer As Integer
Dim bombarray(100) As Integer
Dim bombcount As Integer

Public Sub getnextgrid(direct As Integer, indx As Integer)

 If direct = 1 Then
  getgrid block1(indx).left, block1(indx).top - gridheight
  nextgridx = grx + 1
  nextgridy = gry + 1
 End If
 If direct = 2 Then
  getgrid block1(indx).left + gridwidth, block1(indx).top
  nextgridx = grx + 1
  nextgridy = gry + 1
 End If
 If direct = 3 Then
  getgrid block1(indx).left, block1(indx).top + gridheight
  nextgridx = grx + 1
  nextgridy = gry + 1
 End If
 If direct = 4 Then
  getgrid block1(indx).left - gridwidth, block1(indx).top
  nextgridx = grx + 1
  nextgridy = gry + 1
 End If
 
If (nextgridx < 1) Or (nextgridx > 10) Or (nextgridy < 1) Or (nextgridy > 10) Then
 nextgridx = 0
 nextgridy = 0
End If
 
End Sub

Public Sub shoot(direct As Integer, indx As Integer)

 'Clone enemy missile picture according to indx
 Load missileenemy(indx)
 
 'Set the correct picture for the missile
 'Up
 If direct = 1 Then
  missileenemy(indx).Picture = missup.Picture
 End If
 'Right
 If direct = 2 Then
  missileenemy(indx).Picture = missright.Picture
 End If
 'Down
 If direct = 3 Then
  missileenemy(indx).Picture = missdown.Picture
 End If
 'Left
 If direct = 4 Then
  missileenemy(indx).Picture = missleft.Picture
 End If
 
 'Places initial missile position
 'Up
 If direct = 1 Then missileenemy(indx).Move _
  block1(indx).left + (block1(indx).Width / 2) - (missileenemy(indx).Width / 2), block1(indx).top - _
  missileenemy(indx).Height - 30, missileenemy(indx).Width, _
  missileenemy(indx).Height
 'Right
 If direct = 2 Then missileenemy(indx).Move _
  block1(indx).left + block1(indx).Width + 30, block1(indx).top + _
  (block1(indx).Height / 2) - (missileenemy(indx).Height / 2), missileenemy(indx).Width, _
  missileenemy(indx).Height
 'Down
 If direct = 3 Then missileenemy(indx).Move _
  block1(indx).left + (block1(indx).Width / 2) - (missileenemy(indx).Width / 2), block1(indx).top + _
  block1(indx).Height + 30, missileenemy(indx).Width, _
  missileenemy(indx).Height
 'Left
 If direct = 4 Then missileenemy(indx).Move _
  block1(indx).left - missileenemy(indx).Width - 30, block1(indx).top + _
  (block1(indx).Height / 2) - (missileenemy(indx).Height / 2), missileenemy(indx).Width, _
  missileenemy(indx).Height
  
 missileenemy(indx).Visible = True
 PlaySound App.Path + "\sounds\shoot.wav", 0, SND_ASYNC
 
 Load timmissileenemy(indx)
 timmissileenemy(indx).Enabled = True
 
End Sub
Public Sub gen2()
 
 rand2 = Int((10 * Rnd) + 1)
 
End Sub

Public Sub gen()
 
 rand = Int((5 * Rnd) + 1)
 
End Sub

Public Function readline(filename As String) As String
Dim descr As String

 Open filename For Input As #1
  Input #1, descr
 Close #1
 
 readline = descr
End Function
Public Sub getgrid(X As Integer, Y As Integer)

 grx = X \ block1(0).Width
 gry = Y \ block1(0).Height
 'indexnum = (gry * grdy) + grx

End Sub
Public Sub drawblocks()
 
 'Unloading any previous images
 If blockcount > 0 Then
  For i = 1 To blockcount
   Unload block1(i)
   tankleft.Visible = False
   tankup.Visible = False
   tankright.Visible = False
   tankdown.Visible = False
  Next
  blockcount = 0
 End If
 
 'Set height and width of each block
 block1(0).Width = map.Width / grdx
 block1(0).Height = map.Height / grdy
 gridwidth = block1(0).Width
 gridheight = block1(0).Height
 
 enemyind = 0
 'Set other properties of each block
 For j = 0 To grdy - 1
  For i = 0 To grdx - 1
   If mainarr(i + 1, j + 1) <> 3 Then
    blockcount = blockcount + 1
    Load block1(blockcount)
    block1(blockcount).left = i * block1(0).Width + map.left
    block1(blockcount).top = j * block1(0).Height + map.top
    If mainarr(i + 1, j + 1) = 0 Then block1(blockcount).Picture = emptyblock.Picture
    If mainarr(i + 1, j + 1) = 1 Then block1(blockcount).Picture = wall1.Picture
    If mainarr(i + 1, j + 1) = 2 Then block1(blockcount).Picture = tank1pic.Picture
    block1(blockcount).Visible = True
    'Create array of enemy tank indexes
    If mainarr(i + 1, j + 1) = 2 Then
     block1(blockcount).AutoSize = True
     enemycount = enemycount + 1
     
     '' array of enemy tank indexes only:
     enemyind = enemyind + 1
     ReDim Preserve enemytank(1 To enemycount)
     ReDim Preserve enemyindexes(1 To enemyind)
     enemytank(enemycount) = blockcount
     enemyindexes(enemyind) = blockcount
     '' initialise enemy tracking array
     enemytrack(blockcount).direction = 1
     enemytrack(blockcount).distance = gridwidth
     enemytrack(blockcount).IndexOfTank = blockcount
    End If
    
   Else
    tankup.left = i * block1(0).Width + map.left
    tankup.top = j * block1(0).Height + map.top
    tankup.Visible = True
   End If
  Next
 Next
 
 ''Loads the enemy timers
 For i = 1 To enemyind
  Load frmGame.timEnemy(enemyindexes(i))
 Next
 
End Sub

Private Sub aimethod_Click(Index As Integer)
 aimeth = Index
  
 If aimethod(Index).Picture = check.Picture Then Exit Sub
 
 If aimethod(Index).Picture = uncheck.Picture Then
  aimethod(Index).Picture = check.Picture
  
  If Index = 0 Then
   aimethod(1).Picture = uncheck.Picture
  Else
   aimethod(0).Picture = uncheck.Picture
  End If
  
 Else
  aimethod(Index).Picture = uncheck.Picture
  
  If Index = 1 Then
   aimethod(0).Picture = check.Picture
  Else
   aimethod(1).Picture = check.Picture
  End If

 End If
 
 Text1.SetFocus
 
End Sub


Private Sub avoid_Click()
 Text1.SetFocus
End Sub

Private Sub block1_GotFocus(Index As Integer)
 Text1.SetFocus
End Sub


Private Sub block1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
getgrid block1(Index).left + Int(X) - map.left, block1(Index).top + Int(Y) - map.top
'Text2.Text = str(grx + 1)
'Text3.Text = str(gry + 1)
'Text5.Text = str(Index)
End Sub

Private Sub blowupplayer_Timer()

 bombcount = bombcount + 1
 
 If bombcount = 1 Then
  tankup.Picture = LoadPicture(App.Path + "\pictures\tankexplode.bmp")
  tankright.Picture = LoadPicture(App.Path + "\pictures\tankexplode.bmp")
  tankdown.Picture = LoadPicture(App.Path + "\pictures\tankexplode.bmp")
  tankleft.Picture = LoadPicture(App.Path + "\pictures\tankexplode.bmp")
 End If
 
 If bombcount = 3 Then
  bombcount = 0
  If numlives = 0 Then
   tankup.Visible = False
   tankright.Visible = False
   tankdown.Visible = False
   tankleft.Visible = False
  End If
  tankup.Picture = LoadPicture(App.Path + "\myup2.gif")
  tankright.Picture = LoadPicture(App.Path + "\myright2.gif")
  tankdown.Picture = LoadPicture(App.Path + "\mydown2.gif")
  tankleft.Picture = LoadPicture(App.Path + "\myleft2.gif")
  blowupplayer.Enabled = False
 End If
 
End Sub

Private Sub cmdload_Click()
Dim arrdescr(1 To 223) As String
Dim counter As Integer
Dim j As Integer
Dim myfile As String

 '' Unloads previous enemy timers before loading scene
 For i = 1 To enemyind
  timEnemy(enemyindexes(i)).Enabled = False
  Unload timEnemy(enemyindexes(i))
 Next

myfile = Dir(App.Path + "\maps\*.map")

If myfile = "" Then
 MsgBox "No maps available", vbCritical
 Exit Sub
Else
 counter = counter + 1
 arrfile(counter) = myfile
 arrdescr(counter) = readline(App.Path + "\maps\" + myfile)
End If

Do
 myfile = Dir
 If myfile <> "" Then
  counter = counter + 1
  arrfile(counter) = myfile
  arrdescr(counter) = readline(App.Path + "\maps\" + myfile)
 End If
Loop Until myfile = ""

frmloadmap.lstMaps.Clear
For j = 1 To UBound(arrdescr)
 If arrdescr(j) <> "" Then frmloadmap.lstMaps.AddItem arrdescr(j)
Next

frmloadmap.Show vbModal

End Sub

Private Sub endgame_Click()
 endgame.Enabled = False
 start.Enabled = True
 cmdload.Enabled = True
 
 update.Enabled = False
 gamestarted = False
 Text1.SetFocus
 
 ''Disables the timers of the enemy tanks
 For i = 1 To enemyind
  timEnemy(enemyindexes(i)).Enabled = False
 Next
 
 Text1.SetFocus
 
End Sub

Private Sub exit_Click()
 
 '' Unloads previous enemy timers before exiting
 For i = 1 To enemyind
  timEnemy(enemyindexes(i)).Enabled = False
  Unload timEnemy(enemyindexes(i))
 Next
 
 'Unloading any previous images
 If blockcount > 0 Then
  For i = 1 To blockcount
   Unload block1(i)
   tankleft.Visible = False
   tankup.Visible = False
   tankright.Visible = False
   tankdown.Visible = False
  Next
  blockcount = 0
 End If
 
 End
 
End Sub


Private Sub Form_Activate()
 Text1.Enabled = False
 Text1.Visible = False
 Text1.Visible = True
 Text1.Enabled = True
 Text1.SetFocus
End Sub

Private Sub Form_Load()
 'block1(0).ZOrder (1)
 gamestarted = False
 aimeth = 0

End Sub

Private Sub Label2_Click(Index As Integer)
 Text1.SetFocus
End Sub

Private Sub map_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Text1.SetFocus
End Sub

Private Sub start_Click()
 endgame.Enabled = True
 start.Enabled = False
 cmdload.Enabled = False
 
 gamestarted = True
 update.Enabled = True
 
 Text1.SetFocus
 
 For i = 1 To enemyind
  If enemytank(i) > 0 Then
   timEnemy(enemyindexes(i)).Enabled = True
  End If
 Next
 
End Sub

Private Sub tankdown_Click()
Text1.SetFocus
End Sub

Private Sub tankleft_Click()
Text1.SetFocus
End Sub

Private Sub tankright_Click()
Text1.SetFocus
End Sub

Private Sub tankup_Click()
Text1.SetFocus
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Errorhandler

If gamestarted = False Then Exit Sub
If gameover = True Then Exit Sub

'PlaySound App.Path + "\sounds\tank.wav", 0, SND_ASYNC

If KeyCode = vbKeyUp Then
 
 'check borders
 If tankup.top - 50 < map.top Then Exit Sub
 
 'check for other conflicting objects (i.e. tanks or walls)
 getgrid tankup.left - map.left, tankup.top - map.top - 50
 If (mainarr(grx + 1, gry + 1) <> 0) And (mainarr(grx + 1, gry + 1) <> 3) Then Exit Sub
 getgrid tankup.Width + tankup.left - map.left, tankup.top - map.top - 50
 If (mainarr(grx + 1, gry + 1) <> 0) And (mainarr(grx + 1, gry + 1) <> 3) Then Exit Sub
 
 'move the tank
 tankup.Visible = True
 tankright.Visible = False
 tankdown.Visible = False
 tankleft.Visible = False
 tankup.top = tankup.top - 50
 tankright.top = tankup.top
 tankdown.top = tankup.top
 tankleft.top = tankup.top
 
End If

If KeyCode = vbKeyRight Then
  
 'Alows more 'freedom' for the tanks
 tankright.top = tankup.top
 tankright.left = tankup.left
 
 'check borders
 If tankright.left + tankright.Width + 50 > map.left + map.Width Then Exit Sub

 'check for other conflicting objects (i.e. tanks or walls)
 getgrid tankright.Width + tankright.left - map.left + 50, tankright.top - map.top
 If (mainarr(grx + 1, gry + 1) <> 0) And (mainarr(grx + 1, gry + 1) <> 3) Then Exit Sub
 getgrid tankright.Width + tankright.left - map.left + 50, tankright.top - map.top + tankright.Height
 If (mainarr(grx + 1, gry + 1) <> 0) And (mainarr(grx + 1, gry + 1) <> 3) Then Exit Sub

 tankup.Visible = False
 tankleft.Visible = False
 tankright.Visible = True
 tankdown.Visible = False
 tankright.left = tankright.left + 50
 tankup.left = tankright.left
 tankleft.left = tankright.left
 tankdown.left = tankright.left
End If

If KeyCode = vbKeyLeft Then
 
 tankleft.top = tankup.top
 tankleft.left = tankup.left
 'check borders
 If tankleft.left - 50 < map.left Then Exit Sub
 
 'check for other conflicting objects (i.e. tanks or walls)
 getgrid tankleft.left - map.left - 50, tankleft.top - map.top
 If (mainarr(grx + 1, gry + 1) <> 0) And (mainarr(grx + 1, gry + 1) <> 3) Then Exit Sub
 getgrid tankleft.left - map.left - 50, tankleft.top - map.top + tankleft.Height
 If (mainarr(grx + 1, gry + 1) <> 0) And (mainarr(grx + 1, gry + 1) <> 3) Then Exit Sub

 tankup.Visible = False
 tankright.Visible = False
 tankleft.Visible = True
 tankdown.Visible = False
 tankleft.left = tankleft.left - 50
 tankup.left = tankleft.left
 tankright.left = tankleft.left
 tankdown.left = tankleft.left

End If

If KeyCode = vbKeyDown Then
 
 tankdown.top = tankup.top
 tankdown.left = tankup.left
 'check borders
 If tankdown.top + tankdown.Height + 50 > map.top + map.Height Then Exit Sub

 'check for other conflicting objects (i.e. tanks or walls)
 getgrid tankdown.left - map.left, tankdown.top - map.top + 50 + tankdown.Height
 If (mainarr(grx + 1, gry + 1) <> 0) And (mainarr(grx + 1, gry + 1) <> 3) Then Exit Sub
 getgrid tankdown.Width + tankdown.left - map.left, tankdown.top - map.top + 50 + tankdown.Height
 If (mainarr(grx + 1, gry + 1) <> 0) And (mainarr(grx + 1, gry + 1) <> 3) Then Exit Sub

 tankup.Visible = False
 tankright.Visible = False
 tankleft.Visible = False
 tankdown.Visible = True
 tankdown.top = tankdown.top + 50
 tankleft.top = tankdown.top
 tankright.top = tankdown.top
 tankup.top = tankdown.top
End If

If KeyCode = 17 And timplayer.Enabled = False Then
 If tankup.Visible = True Then
  dirplayer = 1
  missileplayer.Picture = missup.Picture
 End If
 If tankright.Visible = True Then
  dirplayer = 2
  missileplayer.Picture = missright.Picture
 End If
 If tankdown.Visible = True Then
  dirplayer = 3
  missileplayer.Picture = missdown.Picture
 End If
 If tankleft.Visible = True Then
  dirplayer = 4
  missileplayer.Picture = missleft.Picture
 End If
 
 'places initial missile position
 If dirplayer = 1 Then missileplayer.Move _
 tankup.left + (tankup.Width / 2) - (missileplayer.Width / 2), tankup.top - _
 missileplayer.Height - 30, missileplayer.Width, _
 missileplayer.Height
 If dirplayer = 2 Then missileplayer.Move _
 tankright.left + tankright.Width + 30, tankright.top + _
 (tankright.Height / 2) - (missileplayer.Height / 2), missileplayer.Width, _
 missileplayer.Height
 If dirplayer = 3 Then missileplayer.Move _
 tankdown.left + (tankdown.Width / 2) - (missileplayer.Width / 2), tankdown.top + _
 tankdown.Height + 30, missileplayer.Width, _
 missileplayer.Height
 If dirplayer = 4 Then missileplayer.Move _
 tankup.left - missileplayer.Width - 30, tankleft.top + _
 (tankleft.Height / 2) - (missileplayer.Height / 2), missileplayer.Width, _
 missileplayer.Height
 missileplayer.Visible = True
 PlaySound App.Path + "\sounds\shoot.wav", 0, SND_ASYNC
 timplayer.Enabled = True
End If


Errorhandler:
Exit Sub

End Sub

Private Sub timBomb_Timer(Index As Integer)
 
 bombarray(Index) = bombarray(Index) + 1
 
 If bombarray(Index) = 1 Then block1(Index).Picture = LoadPicture(App.Path + "\pictures\tankexplode.bmp")
 
 If bombarray(Index) = 10 Then
  bombarray(Index) = 0
  block1(Index).Picture = emptyblock.Picture
  block1(Index).Visible = False
  Unload timBomb(Index)
 End If
 
End Sub

Private Sub timEnemy_Timer(Index As Integer)
Dim r As Integer
Dim prevleft, prevtop As Integer
Dim distancex, distancey As Integer
Dim done As Boolean

On Error GoTo Errorhandler
 
  
 ''Generate a new direction (if necessary)
 If enemytrack(Index).distance > 0 Then
  If (enemytrack(Index).direction = 1) Or (enemytrack(Index).direction = 3) Then
   enemytrack(Index).distance = enemytrack(Index).distance - (gridheight / 18)
  End If
  If (enemytrack(Index).direction = 2) Or (enemytrack(Index).direction = 4) Then
   enemytrack(Index).distance = enemytrack(Index).distance - (gridwidth / 18)
  End If
 Else 'The tank has reached a new grid position
  enemytrack(Index).indanger = False
  getnextgrid enemytrack(Index).direction, Index
  
  ''Random method:
  If aimethod(0).Picture = check.Picture Then
   If (nextgridx > 0) And (nextgridy > 0) Then
    If mainarr(nextgridx, nextgridy) = 0 Then
    rand = enemytrack(Index).direction
     Else
     gen
    End If
   Else
    gen
   End If
  End If
      
  ''Shortest axis method
  If aimethod(1).Picture = check.Picture Then
   distancex = 0
   distancey = 0
   If tankup.left > block1(Index).left Then distancex = tankup.left - block1(Index).left
   If tankup.left < block1(Index).left Then distancex = block1(Index).left - tankup.left
   If tankup.top > block1(Index).top Then distancey = tankup.top - block1(Index).top
   If tankup.top < block1(Index).top Then distancey = block1(Index).top - tankup.top
   'Check which direction to go
   If distancex <= distancey Then
    If tankup.left > block1(Index).left Then
     rand = 2
    Else
     rand = 4
    End If
   Else
    If tankup.top > block1(Index).top Then
     rand = 3
    Else
     rand = 1
    End If
   End If
   'Check if next block is empty
   r = rand
   getnextgrid r, Index
   If (nextgridx > 0) And (nextgridy > 0) Then
    If mainarr(nextgridx, nextgridy) = 1 Then shoot r, Index
   End If
  End If
      
  If (rand = 1) Or (rand = 3) Then
   enemytrack(Index).distance = gridheight
  End If
  If (rand = 2) Or (rand = 4) Then
   enemytrack(Index).distance = gridwidth
  End If
  enemytrack(Index).direction = rand
 
 End If 'end if of next grid position
 
 If enemytrack(Index).indanger = False Then
 ''Tries to avoid missiles
  If (avoid.Value = 1) And (missileplayer.Visible = True) Then
    
   If missileplayer.Picture = missup.Picture Then
   done = False
    If (missileplayer.left > block1(Index).left) And (missileplayer.left < block1(Index).left + block1(Index).Width) Then
     If (missileplayer.top - 10000 < block1(Index).top) Then
      rand = 4
      enemytrack(Index).distance = gridwidth
      done = True
      enemytrack(Index).indanger = True
      enemytrack(Index).direction = rand
      
     End If
    End If
    If (missileplayer.left + missileplayer.Width > block1(Index).left) And (missileplayer.left + missileplayer.Width < block1(Index).left + block1(Index).Width) Then
     If (missileplayer.top - 10000 < block1(Index).top) Then
      If done = False Then
       rand = 2
       enemytrack(Index).distance = gridwidth
       enemytrack(Index).indanger = True
       enemytrack(Index).direction = rand
       
      Else
       gen2
       If rand2 < 6 Then rand = 2
       If rand2 > 5 Then rand = 4
       enemytrack(Index).distance = gridwidth
       enemytrack(Index).indanger = True
       enemytrack(Index).direction = rand
       
      End If
     End If
    End If
   End If
   
   If missileplayer.Picture = missright.Picture Then
   done = False
    If (missileplayer.top > block1(Index).top) And (missileplayer.top < block1(Index).top + block1(Index).Height) Then
     If (missileplayer.left + 10000 > block1(Index).left + block1(Index).Width) Then
      rand = 1
      enemytrack(Index).distance = gridheight
      done = True
      enemytrack(Index).indanger = True
      enemytrack(Index).direction = rand
      
     End If
    End If
    If (missileplayer.top + missileplayer.Height > block1(Index).top) And (missileplayer.top + missileplayer.Height < block1(Index).top + block1(Index).Height) Then
     If (missileplayer.left + 10000 > block1(Index).left + block1(Index).Width) Then
      If done = False Then
       rand = 3
       enemytrack(Index).distance = gridheight
       enemytrack(Index).indanger = True
       enemytrack(Index).direction = rand
       
      Else
       gen2
       If rand2 < 6 Then rand = 3
       If rand2 > 5 Then rand = 1
       enemytrack(Index).distance = gridheight
       enemytrack(Index).indanger = True
       enemytrack(Index).direction = rand
       
      End If
     End If
    End If
   End If
   
   If missileplayer.Picture = missdown.Picture Then
    done = False
    If (missileplayer.left > block1(Index).left) And (missileplayer.left < block1(Index).left + block1(Index).Width) Then
     If (missileplayer.top + 10000 > block1(Index).top + block1(Index).Height) Then
      rand = 4
      enemytrack(Index).distance = gridwidth
      done = True
      enemytrack(Index).indanger = True
      enemytrack(Index).direction = rand
      
     End If
    End If
    If (missileplayer.left + missileplayer.Width > block1(Index).left) And (missileplayer.left + missileplayer.Width < block1(Index).left + block1(Index).Width) Then
     If (missileplayer.top + 10000 > block1(Index).top + block1(Index).Height) Then
      If done = False Then
       rand = 2
       enemytrack(Index).distance = gridwidth
       enemytrack(Index).indanger = True
       enemytrack(Index).direction = rand
       
      Else
       gen2
       If rand2 < 6 Then rand = 2
       If rand2 > 5 Then rand = 4
       enemytrack(Index).distance = gridwidth
       enemytrack(Index).indanger = True
       enemytrack(Index).direction = rand
       
      End If
     End If
    End If
   End If
   
   If missileplayer.Picture = missleft.Picture Then
    done = False
    If (missileplayer.top > block1(Index).top) And (missileplayer.top < block1(Index).top + block1(Index).Height) Then
     If (missileplayer.left - 10000 < block1(Index).left) Then
      rand = 1
      enemytrack(Index).distance = gridheight
      done = True
      enemytrack(Index).indanger = True
      enemytrack(Index).direction = rand
      
     End If
    End If
    If (missileplayer.top + missileplayer.Height > block1(Index).top) And (missileplayer.top + missileplayer.Height < block1(Index).top + block1(Index).Height) Then
     If (missileplayer.left - 10000 < block1(Index).left) Then
      If done = False Then
       rand = 3
       enemytrack(Index).distance = gridheight
       enemytrack(Index).indanger = True
       enemytrack(Index).direction = rand
       
      Else
       gen2
       If rand2 < 6 Then rand = 3
       If rand2 > 5 Then rand = 1
       enemytrack(Index).distance = gridheight
       enemytrack(Index).indanger = True
       enemytrack(Index).direction = rand
       
      End If
     End If
    End If
   End If
  End If
 End If
  
 ''Checks tanks that are frozen
 If (block1(Index).left <> enemytrack(Index).left) Or (block1(Index).top <> enemytrack(Index).top) Then
  enemytrack(Index).left = block1(Index).left
  enemytrack(Index).top = block1(Index).top
  enemytrack(Index).timeout = 0
 End If
 If (block1(Index).left = enemytrack(Index).left) And (block1(Index).top = enemytrack(Index).top) Then
  enemytrack(Index).timeout = enemytrack(Index).timeout + 1
  If enemytrack(Index).timeout = 2 Then
   gen
   enemytrack(Index).direction = rand
   If (rand = 1) Or (rand = 3) Then
    enemytrack(Index).distance = gridheight
   End If
   If (rand = 2) Or (rand = 4) Then
    enemytrack(Index).distance = gridwidth
   End If
   enemytrack(Index).timeout = 0
  End If
 End If
 
 ''Searches for the players tank
 If enemytrack(Index).indanger = False Then
  If (block1(Index).top - 49 < tankup.top) And (block1(Index).top + 49 > tankup.top) Then
   If block1(Index).left < tankup.left Then
    enemytrack(Index).direction = 2
    enemytrack(Index).distance = gridwidth
    If block1(Index).Picture <> enemyright.Picture Then
     block1(Index).left = block1(Index).left - (enemyup.Height - enemyup.Width)
    End If
    block1(Index).Picture = enemyright.Picture
    shoot 2, Index
    Exit Sub
   Else
    enemytrack(Index).direction = 4
    enemytrack(Index).distance = gridwidth
    If block1(Index).Picture <> enemyleft.Picture Then
     block1(Index).left = block1(Index).left + (enemyup.Height - enemyup.Width)
    End If
    block1(Index).Picture = enemyleft.Picture
    shoot 4, Index
    Exit Sub
   End If
  End If
  If (block1(Index).left - 49 < tankup.left) And (block1(Index).left + 49 > tankup.left) Then
   If block1(Index).top < tankup.top Then
    enemytrack(Index).direction = 3
    enemytrack(Index).distance = gridheight
    If block1(Index).Picture <> enemydown.Picture Then
     block1(Index).top = block1(Index).top - (enemyup.Height - enemyup.Width)
    End If
    block1(Index).Picture = enemydown.Picture
    shoot 3, Index
    Exit Sub
   Else
    enemytrack(Index).direction = 1
    enemytrack(Index).distance = gridheight
    If block1(Index).Picture <> enemyup.Picture Then
     block1(Index).top = block1(Index).top + (enemyup.Height - enemyup.Width)
    End If
    block1(Index).Picture = enemyup.Picture
    shoot 1, Index
    Exit Sub
   End If
  End If
 End If
 
 
 ''Move the enemy tanks:
 
 'Up
 If enemytrack(Index).direction = 1 Then
   'block1(Index).Picture = enemyup.Picture
   
   'check borders
   If block1(Index).top - 50 < map.top Then Exit Sub
  
   'check for other conflicting objects (i.e. tanks, walls, or player)
   getgrid block1(Index).left - map.left, block1(Index).top - map.top - 50
   If (mainarr(grx + 1, gry + 1) <> 0) And (mainarr(grx + 1, gry + 1) <> 2) Then Exit Sub
   getgrid block1(Index).Width + block1(Index).left - map.left, block1(Index).top - map.top - 50
   If (mainarr(grx + 1, gry + 1) <> 0) And (mainarr(grx + 1, gry + 1) <> 2) Then Exit Sub
  
   'move the tank
   block1(Index).Picture = enemyup.Picture
   block1(Index).ZOrder (0)
   block1(Index).top = block1(Index).top - (gridheight / 18)
 End If
 
 'Right
 If enemytrack(Index).direction = 2 Then
   'block1(Index).Picture = enemyright.Picture
   
   'check borders
   If block1(Index).left + block1(Index).Width + 50 > map.left + map.Width Then Exit Sub
  
   'check for other conflicting objects (i.e. tanks, walls, or player)
   getgrid block1(Index).left + block1(Index).Width - map.left + 50, block1(Index).top - map.top
   If (mainarr(grx + 1, gry + 1) <> 0) And (mainarr(grx + 1, gry + 1) <> 2) Then Exit Sub
   getgrid block1(Index).Width + block1(Index).left - map.left + 50, block1(Index).top - map.top + block1(Index).Height
   If (mainarr(grx + 1, gry + 1) <> 0) And (mainarr(grx + 1, gry + 1) <> 2) Then Exit Sub
  
   'move the tank
   block1(Index).Picture = enemyright.Picture
   block1(Index).ZOrder (0)
   block1(Index).left = block1(Index).left + (gridwidth / 18)
 End If
 
 'Down
 If enemytrack(Index).direction = 3 Then
   'block1(Index).Picture = enemydown.Picture
   
   'check borders
   If block1(Index).top + block1(Index).Height + 50 > map.top + map.Height Then Exit Sub
  
   'check for other conflicting objects (i.e. tanks, walls, or player)
   getgrid block1(Index).left - map.left, block1(Index).top - map.top + 50 + block1(Index).Height
   If (mainarr(grx + 1, gry + 1) <> 0) And (mainarr(grx + 1, gry + 1) <> 2) Then Exit Sub
   getgrid block1(Index).Width + block1(Index).left - map.left, block1(Index).top - map.top + 50 + block1(Index).Height
   If (mainarr(grx + 1, gry + 1) <> 0) And (mainarr(grx + 1, gry + 1) <> 2) Then Exit Sub
  
   'move the tank
   block1(Index).Picture = enemydown.Picture
   block1(Index).ZOrder (0)
   block1(Index).top = block1(Index).top + (gridheight / 18)
 End If
 
 'Left
 If enemytrack(Index).direction = 4 Then
   'block1(Index).Picture = enemyleft.Picture
   
   'check borders
   If block1(Index).left - 50 < map.left Then Exit Sub
  
   'check for other conflicting objects (i.e. tanks, walls, or player)
   getgrid block1(Index).left - map.left - 50, block1(Index).top - map.top
   If (mainarr(grx + 1, gry + 1) <> 0) And (mainarr(grx + 1, gry + 1) <> 2) Then Exit Sub
   getgrid block1(Index).left - map.left - 50, block1(Index).top - map.top + block1(Index).Height
   If (mainarr(grx + 1, gry + 1) <> 0) And (mainarr(grx + 1, gry + 1) <> 2) Then Exit Sub
  
   'move the tank
   block1(Index).Picture = enemyleft.Picture
   block1(Index).ZOrder (0)
   block1(Index).left = block1(Index).left - (gridwidth / 18)
 End If
 
Errorhandler:
Exit Sub
  
End Sub

Private Sub timmissileenemy_Timer(Index As Integer)
Dim j As Integer
Dim avoidfor As Boolean

On Error GoTo Errorhandler
missileenemy(Index).ZOrder (0)

''Check borders
'Up
If missileenemy(Index).Picture = missup.Picture Then
 If missileenemy(Index).top - 50 < map.top Then
  missileenemy(Index).Visible = False
  timmissileenemy(Index).Enabled = False
  Unload missileenemy(Index)
  Unload timmissileenemy(Index)
  Exit Sub
 End If
End If
'Right
If missileenemy(Index).Picture = missright.Picture Then
 If missileenemy(Index).left + missileenemy(Index).Width + 50 > map.left + map.Width Then
  missileenemy(Index).Visible = False
  timmissileenemy(Index).Enabled = False
  Unload missileenemy(Index)
  Unload timmissileenemy(Index)
  Exit Sub
 End If
End If
'Down
If missileenemy(Index).Picture = missdown.Picture Then
 If missileenemy(Index).top + missileenemy(Index).Height + 50 > map.top + map.Height Then
  missileenemy(Index).Visible = False
  timmissileenemy(Index).Enabled = False
  Unload missileenemy(Index)
  Unload timmissileenemy(Index)
  Exit Sub
 End If
End If
'Left
If missileenemy(Index).Picture = missleft.Picture Then
 If missileenemy(Index).left - 50 < map.left Then
  missileenemy(Index).Visible = False
  timmissileenemy(Index).Enabled = False
  Unload missileenemy(Index)
  Unload timmissileenemy(Index)
  Exit Sub
 End If
End If

'Gets missile's grid block
getgrid missileenemy(Index).left - map.left, missileenemy(Index).top - map.top
If mainarr(grx + 1, gry + 1) = 0 Or mainarr(grx + 1, gry + 1) = 2 Then
 getgrid missileenemy(Index).left + missileenemy(Index).Width - map.left, missileenemy(Index).top + missileenemy(Index).Height - map.top
 If mainarr(grx + 1, gry + 1) = 0 Or mainarr(grx + 1, gry + 1) = 2 Then
  avoidfor = True
 End If
End If

If avoidfor = False Then
For j = 1 To blockcount
 If (block1(j).Picture = wall1.Picture) Then
 
 'top left corner
  If (missileenemy(Index).top > block1(j).top) And (missileenemy(Index).top < block1(j).top + block1(j).Height) And (missileenemy(Index).left > block1(j).left) And (missileenemy(Index).left < block1(j).left + block1(j).Width) Then
    missileenemy(Index).Visible = False
    timmissileenemy(Index).Enabled = False
    PlaySound App.Path + "\sounds\explode.wav", 0, SND_ASYNC
    
    getgrid block1(j).left, block1(j).top
    mainarr(grx + 1, gry + 1) = 0
    
    'display the bomb
    Load timBomb(j)
    timBomb(j).Enabled = True
    
    Unload missileenemy(Index)
    Unload timmissileenemy(Index)
    
    'block1(j).Visible = False
    Exit Sub
  End If
 
 'top right corner
  If (missileenemy(Index).top > block1(j).top) And (missileenemy(Index).top < block1(j).top + block1(j).Height) And (missileenemy(Index).left + missileenemy(Index).Width > block1(j).left) And (missileenemy(Index).left + missileenemy(Index).Width < block1(j).left + block1(j).Width) Then
    missileenemy(Index).Visible = False
    timmissileenemy(Index).Enabled = False
    PlaySound App.Path + "\sounds\explode.wav", 0, SND_ASYNC
    
    getgrid block1(j).left, block1(j).top
    mainarr(grx + 1, gry + 1) = 0
    'display bomb
    
    Load timBomb(j)
    timBomb(j).Enabled = True
    
    Unload missileenemy(Index)
    Unload timmissileenemy(Index)
    
    'block1(j).Visible = False
    Exit Sub
  End If
 
 'bottom left corner
  If (missileenemy(Index).top + missileenemy(Index).Height > block1(j).top) And (missileenemy(Index).top + missileenemy(Index).Height < block1(j).top + block1(j).Height) And (missileenemy(Index).left > block1(j).left) And (missileenemy(Index).left < block1(j).left + block1(j).Width) Then
    missileenemy(Index).Visible = False
    timmissileenemy(Index).Enabled = False
    PlaySound App.Path + "\sounds\explode.wav", 0, SND_ASYNC
    
    getgrid block1(j).left, block1(j).top
    mainarr(grx + 1, gry + 1) = 0
    'display bomb
    
    Load timBomb(j)
    timBomb(j).Enabled = True
    
    Unload missileenemy(Index)
    Unload timmissileenemy(Index)
    
    'block1(j).Visible = False
    Exit Sub
  End If
  
 'bottom right corner
  If (missileenemy(Index).top + missileenemy(Index).Height > block1(j).top) And (missileenemy(Index).top + missileenemy(Index).Height < block1(j).top + block1(j).Height) And (missileenemy(Index).left + missileenemy(Index).Width > block1(j).left) And (missileenemy(Index).left + missileenemy(Index).Width < block1(j).left + block1(j).Width) Then
    missileenemy(Index).Visible = False
    timmissileenemy(Index).Enabled = False
    PlaySound App.Path + "\sounds\explode.wav", 0, SND_ASYNC
    
    getgrid block1(j).left, block1(j).top
    mainarr(grx + 1, gry + 1) = 0
    'display bomb
    
    Load timBomb(j)
    timBomb(j).Enabled = True
    
    Unload missileenemy(Index)
    Unload timmissileenemy(Index)

    'block1(j).Visible = False
    Exit Sub
  End If
 End If
Next
End If


 ''Check if the missile hits the player
 If avoidfor = False Then
  numlives = numlives - 1
  piclives(numlives + 1).Visible = False
  missileenemy(Index).Visible = False
  timmissileenemy(Index).Enabled = False
  PlaySound App.Path + "\sounds\explode.wav", 0, SND_ASYNC
  bombcount = 0
  blowupplayer.Enabled = True
  'If no more lives remain, make the tank dissapear
  If numlives = 0 Then
   gameover = True
   getgrid missileenemy(Index).left - map.left, missileenemy(Index).top - map.top
   If mainarr(grx + 1, gry + 1) = 3 Then mainarr(grx + 1, gry + 1) = 0
   getgrid missileenemy(Index).left + missileenemy(Index).Width - map.left, missileenemy(Index).top + missileenemy(Index).Height - map.top
   If mainarr(grx + 1, gry + 1) = 3 Then mainarr(grx + 1, gry + 1) = 0
   tankup.Visible = False
   tankright.Visible = False
   tankdown.Visible = False
   tankleft.Visible = False
   endgame_Click
   start.Enabled = False
   MsgBox "Game Over!", , "Message"
  End If
  Unload missileenemy(Index)
  Unload timmissileenemy(Index)
  Exit Sub
 End If
 
 
 'move the missile in the appropriate direction
 If missileenemy(Index).Picture = missup.Picture Then missileenemy(Index).top = missileenemy(Index).top - 300
 If missileenemy(Index).Picture = missright.Picture Then missileenemy(Index).left = missileenemy(Index).left + 300
 If missileenemy(Index).Picture = missdown.Picture Then missileenemy(Index).top = missileenemy(Index).top + 300
 If missileenemy(Index).Picture = missleft.Picture Then missileenemy(Index).left = missileenemy(Index).left - 300

Errorhandler:
 Exit Sub
 missileenemy(Index).Visible = False
 timmissileenemy(Index).Enabled = False
 Unload missileenemy(Index)
 Unload timmissileenemy(Index)
 
End Sub

Private Sub timplayer_Timer()
Dim j As Integer
Dim avoidfor As Boolean

On Error GoTo Errorhandler
missileplayer.ZOrder (0)

'Check borders
If dirplayer = 1 Then
 If missileplayer.top - 50 < map.top Then
  missileplayer.Visible = False
  timplayer.Enabled = False
  Exit Sub
 End If
End If
If dirplayer = 2 Then
 If missileplayer.left + missileplayer.Width + 50 > map.left + map.Width Then
  missileplayer.Visible = False
  timplayer.Enabled = False
  Exit Sub
 End If
End If
If dirplayer = 3 Then
 If missileplayer.top + missileplayer.Height + 50 > map.top + map.Height Then
  missileplayer.Visible = False
  timplayer.Enabled = False
  Exit Sub
 End If
End If
If dirplayer = 4 Then
 If missileplayer.left - 50 < map.left Then
  missileplayer.Visible = False
  timplayer.Enabled = False
  Exit Sub
 End If
End If

'Gets missile's grid block
getgrid missileplayer.left - map.left, missileplayer.top - map.top
If mainarr(grx + 1, gry + 1) = 0 Or mainarr(grx + 1, gry + 1) = 3 Then
 getgrid missileplayer.left + missileplayer.Width - map.left, missileplayer.top + missileplayer.Height - map.top
 If mainarr(grx + 1, gry + 1) = 0 Or mainarr(grx + 1, gry + 1) = 3 Then
  avoidfor = True
 End If
End If

If avoidfor = False Then
For j = 1 To blockcount
 If block1(j).Picture <> emptyblock.Picture Then
 
 'top left corner
  If (missileplayer.top > block1(j).top) And (missileplayer.top < block1(j).top + block1(j).Height) And (missileplayer.left > block1(j).left) And (missileplayer.left < block1(j).left + block1(j).Width) Then
    missileplayer.Visible = False
    timplayer.Enabled = False
    PlaySound App.Path + "\sounds\explode.wav", 0, SND_ASYNC
    For i = 1 To enemycount
     If enemytank(i) = j Then
      enemytank(i) = -1
      enemytrack(j).IndexOfTank = 0
      timEnemy(j).Enabled = False
     End If
    Next
    getgrid block1(j).left, block1(j).top
    mainarr(grx + 1, gry + 1) = 0
    'display the bomb
    
    Load timBomb(j)
    timBomb(j).Enabled = True
    'block1(j).Visible = False
    Exit Sub
  End If
 
 'top right corner
  If (missileplayer.top > block1(j).top) And (missileplayer.top < block1(j).top + block1(j).Height) And (missileplayer.left + missileplayer.Width > block1(j).left) And (missileplayer.left + missileplayer.Width < block1(j).left + block1(j).Width) Then
    missileplayer.Visible = False
    timplayer.Enabled = False
    PlaySound App.Path + "\sounds\explode.wav", 0, SND_ASYNC
    For i = 1 To enemycount
     If enemytank(i) = j Then
      enemytank(i) = -1
      enemytrack(j).IndexOfTank = 0
      timEnemy(j).Enabled = False
     End If
    Next
    getgrid block1(j).left, block1(j).top
    mainarr(grx + 1, gry + 1) = 0
    'display bomb
    
    Load timBomb(j)
    timBomb(j).Enabled = True
    'block1(j).Visible = False
    Exit Sub
  End If
 
 'bottom left corner
  If (missileplayer.top + missileplayer.Height > block1(j).top) And (missileplayer.top + missileplayer.Height < block1(j).top + block1(j).Height) And (missileplayer.left > block1(j).left) And (missileplayer.left < block1(j).left + block1(j).Width) Then
    missileplayer.Visible = False
    timplayer.Enabled = False
    PlaySound App.Path + "\sounds\explode.wav", 0, SND_ASYNC
    For i = 1 To enemycount
     If enemytank(i) = j Then
      enemytank(i) = -1
      enemytrack(j).IndexOfTank = 0
      timEnemy(j).Enabled = False
     End If
    Next
    getgrid block1(j).left, block1(j).top
    mainarr(grx + 1, gry + 1) = 0
    'display bomb
    
    Load timBomb(j)
    timBomb(j).Enabled = True
    'block1(j).Visible = False
    Exit Sub
  End If
  
 'bottom right corner
  If (missileplayer.top + missileplayer.Height > block1(j).top) And (missileplayer.top + missileplayer.Height < block1(j).top + block1(j).Height) And (missileplayer.left + missileplayer.Width > block1(j).left) And (missileplayer.left + missileplayer.Width < block1(j).left + block1(j).Width) Then
    missileplayer.Visible = False
    timplayer.Enabled = False
    PlaySound App.Path + "\sounds\explode.wav", 0, SND_ASYNC
    For i = 1 To enemycount
     If enemytank(i) = j Then
      enemytank(i) = -1
      enemytrack(j).IndexOfTank = 0
      timEnemy(j).Enabled = False
     End If
    Next
    getgrid block1(j).left, block1(j).top
    mainarr(grx + 1, gry + 1) = 0
    'display bomb
    
    Load timBomb(j)
    timBomb(j).Enabled = True
    'block1(j).Visible = False
    Exit Sub
  End If
 End If
Next
End If

 'move the missile in the appropriate direction
 If dirplayer = 1 Then missileplayer.top = missileplayer.top - 200
 If dirplayer = 2 Then missileplayer.left = missileplayer.left + 200
 If dirplayer = 3 Then missileplayer.top = missileplayer.top + 200
 If dirplayer = 4 Then missileplayer.left = missileplayer.left - 200

Errorhandler:
 Exit Sub
 missileplayer.Visible = False
 timplayer.Enabled = False
 
End Sub

Private Sub update_Timer()
On Error GoTo Errorhandler

 'empties 'non-wall' blocks
 For j = 1 To grdy
  For i = 1 To grdx
   If mainarr(i, j) > 1 Then mainarr(i, j) = 0
  Next
 Next
 
 'updates the block array (enemy tanks)
 For i = 1 To enemycount
  If enemytank(i) > -1 Then
   getgrid block1(enemytank(i)).left + (block1(enemytank(i)).Width / 2) - map.left, block1(enemytank(i)).top + (block1(enemytank(i)).Height / 2) - map.top
   mainarr(grx + 1, gry + 1) = 2
  End If
 Next
  
 'updates the block array (player's tank)
 If tankright.Visible = True Then
  getgrid tankright.left + (tankright.Width / 2) - map.left, tankright.top + (tankright.Height / 2) - map.TabIndex
 End If
 If tankdown.Visible = True Then
  getgrid tankdown.left + (tankdown.Width / 2) - map.left, tankdown.top + (tankdown.Height / 2) - map.top
 End If
 If tankleft.Visible = True Then
  getgrid tankleft.left + (tankleft.Width / 2) - map.left, tankleft.top + (tankleft.Height / 2) - map.top
 End If
 If tankup.Visible = True Then
  getgrid tankup.left + (tankup.Width / 2) - map.left, tankup.top + (tankup.Height / 2) - map.top
 End If
 mainarr(grx + 1, gry + 1) = 3
 'Shape1.Move (grx * block1(0).Width) + map.Left, (gry * block1(0).Height) + map.Top, Shape1.Width, Shape1.Height
 'Shape1.Visible = True
 'Shape1.ZOrder (0)
 
Errorhandler:
 Exit Sub
 
End Sub
