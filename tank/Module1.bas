Attribute VB_Name = "TankModule"
Option Explicit

Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)


'PlaySound Constants
Public Const SND_ASYNC = &H1
Public Const SND_LOOP = &H8
Public Const SND_NOSTOP = &H10


Public arrfile(1 To 223) As String
Public mainarr() As Integer
Public blockcount As Integer
Public grdx, grdy As Integer
Public indexnum As Integer
Public grx, gry As Integer
Public enemytank() As Integer
Public enemyindexes() As Integer
Public enemycount As Integer
Public enemyind As Integer
Public rand, rand2 As Integer
Public numlives As Integer
Public gameover As Boolean
Public gridwidth, gridheight As Integer
Public nextgridx, nextgridy As Integer
Public aimeth As Integer

'Enemy tracking array of records
Public Type type1
 distance As Integer
 direction As Integer
 IndexOfTank As Integer
 left As Integer
 top As Integer
 timeout As Integer
 indanger As Boolean
End Type

Public enemytrack() As type1
 
