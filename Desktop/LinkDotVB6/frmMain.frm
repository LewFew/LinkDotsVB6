VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LINK DOTS V2"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   9930
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrClock 
      Interval        =   10
      Left            =   8880
      Top             =   2160
   End
   Begin VB.PictureBox picDisplay 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   5175
      Left            =   0
      ScaleHeight     =   5175
      ScaleWidth      =   7215
      TabIndex        =   0
      Top             =   0
      Width           =   7215
   End
   Begin VB.Timer tmrLoop 
      Interval        =   1
      Left            =   9000
      Top             =   960
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Author: William Chan
'Date: May 17th, 2019
'Purpose: ICS4U Culminating Assignment

Dim KeyDown() As Boolean
Dim MouseDown() As Boolean

Public Function IsKeyDownMAIN(ByVal Key As Integer) As Boolean
    IsKeyDownMAIN = KeyDown(Key)
End Function

Public Function IsMouseDownMAIN(ByVal Button As Integer) As Boolean
    IsMouseDownMAIN = MouseDown(Button)
End Function

Public Sub ResetKeys()
    ReDim KeyDown(0 To 1000)
    ReDim MouseDown(0 To 3)
End Sub

Private Sub Form_GotFocus()
    ResetKeys
End Sub

Private Sub Form_Load()
    Running = True
    
    ReDim KeyDown(0 To 1000)
    ReDim MouseDown(0 To 3)
    
    picDisplay.Top = 0
    picDisplay.Left = 0
    picDisplay.Width = frmMain.Width
    picDisplay.Height = frmMain.Height

    Init
    LoadNewLevel
    
    LevelsCompleted = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload frmImages
    Unload frmMain
End Sub

Private Sub picDisplay_GotFocus()
    ResetKeys
End Sub

Private Sub picDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyDown(KeyCode) = True
End Sub

Private Sub picDisplay_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyDown(KeyCode) = False
End Sub

Private Sub picDisplay_LostFocus()
    ResetKeys
End Sub

Private Sub picDisplay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseDown(Button) = True
    MouseX = X
    MouseY = Y
End Sub

Private Sub picDisplay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseX = X
    MouseY = Y
    If (AbsoluteCorrectX(X) - PlayerX <> 0) Then
        MouseAngle = ToDegrees(Atn(-(AbsoluteCorrectY(Y) - (PlayerY + ToTwips(PLAYERHEIGHT / 2))) / (AbsoluteCorrectX(X) - (PlayerX + PLAYERWIDTH / 2))))
    Else
        MouseAngle = 0
    End If
    
    If (AbsoluteCorrectX(X) - PlayerX < 0) Then
        MouseAngle = MouseAngle + 180
    End If
    
    If (AbsoluteCorrectY(Y) - PlayerY > 0 And AbsoluteCorrectX(X) - PlayerX > 0) Then
        MouseAngle = MouseAngle + 360
    End If
End Sub

Private Sub picDisplay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseDown(Button) = False
    MouseX = X
    MouseY = Y
End Sub

Private Sub tmrClock_Timer()
    Clock = Round(Clock + 1, 0)
End Sub

Private Sub tmrLoop_Timer()
    If (Running) Then
        picDisplay.Cls
        Update
        Render
    Else
        Form_QueryUnload 0, 0
    End If
End Sub
