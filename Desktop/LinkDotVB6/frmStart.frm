VERSION 5.00
Begin VB.Form frmStart 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LINK DOTS - Menu (Version 2)"
   ClientHeight    =   4545
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   8190
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   3720
      Width           =   7695
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   3000
      Width           =   7695
   End
   Begin VB.Label lblTitle 
      Caption         =   "LINK DOTS"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808000&
      FillColor       =   &H00FFFF80&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   5
      Left            =   7200
      Shape           =   3  'Circle
      Top             =   1440
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808000&
      FillColor       =   &H00FFFF80&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   4
      Left            =   5040
      Shape           =   3  'Circle
      Top             =   720
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808000&
      FillColor       =   &H00FFFF80&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   3
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   2520
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808000&
      FillColor       =   &H00FFFF80&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   2
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808000&
      FillColor       =   &H00FFFF80&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   1
      Left            =   1680
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808000&
      FillColor       =   &H00FFFF80&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   360
      Shape           =   3  'Circle
      Top             =   240
      Width           =   375
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      X1              =   1800
      X2              =   3240
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      X1              =   3240
      X2              =   4800
      Y1              =   1680
      Y2              =   2640
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      X1              =   4800
      X2              =   5160
      Y1              =   2640
      Y2              =   840
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      X1              =   5160
      X2              =   7440
      Y1              =   840
      Y2              =   1560
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      X1              =   7320
      X2              =   480
      Y1              =   1560
      Y2              =   360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      X1              =   600
      X2              =   1800
      Y1              =   480
      Y2              =   1680
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuTutorial 
         Caption         =   "How To Play"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuHistory 
      Caption         =   "History"
      Begin VB.Menu mnuLeaderboard 
         Caption         =   "Leaderboard"
      End
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Author: William Chan
'Date: May 17th, 2019
'Purpose: ICS4U Culminating Assignment

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdStart_Click()
    Randomize
    frmMain.Show
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload frmMain
    Unload frmImages
    Unload frmAbout
    Unload frmTutorial
    Unload frmScores
    Unload frmStart
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuLeaderboard_Click()
    frmScores.Show
    frmScores.ReadLeaderboard
    frmScores.ShowGrid
End Sub

Private Sub mnuTutorial_Click()
    frmTutorial.Show
End Sub
