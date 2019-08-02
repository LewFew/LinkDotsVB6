VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmScores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Link Dots - High Scores"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   4440
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid grdScores 
      Height          =   1550
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   2725
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblTitle 
      Caption         =   "Link Dots Leaderboard"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmScores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const WIDTHOFFSETX = 55
Const HEIGHTOFFSETY = 100
Const MAXENTRIES = 5
Const RANKWIDTH = 300

Dim NumRecs As Integer

Dim Scores(1 To MAXENTRIES) As RecordScore

Public Function GetMaxEntries() As Integer
    GetMaxEntries = MAXENTRIES
End Function

Public Function GetLevelsCompleted(ByVal Index As Integer) As Integer
    GetLevelsCompleted = Scores(Index).LevelsComplete
End Function

Public Sub AddEntry(ByVal Name As String, ByVal CompletedLevels As Integer)
    Dim Rec As RecordScore
    Rec.Name = Name
    Rec.LevelsComplete = CompletedLevels
    If (NumRecs < MAXENTRIES) Then
        Scores(NumRecs + 1) = Rec
        NumRecs = NumRecs + 1
    Else
        Scores(NumRecs) = Rec
    End If
    SaveLeaderboard
    ShowGrid
End Sub

Public Sub SaveLeaderboard()
    Dim FileName As String
    Dim X As Integer
    Dim RecLength As Integer
    Dim DummyRecord As RecordScore
    
    RecLength = Len(DummyRecord)
    
    FileName = "\HighScores.rec"
    
    X = 0
    
    On Error GoTo ErrorHandler
    Kill FileName
        
    Open App.Path & FileName For Random As #1 Len = RecLength
        For X = 1 To NumRecs
            Put #1, X, Scores(X)
        Next X
    Close #1
    
    Exit Sub

ErrorHandler:
    Resume Next
End Sub

Public Sub ReadLeaderboard()
    Dim I As Integer
    Dim J As Integer
    Dim X As Integer
    Dim FName As String
    Dim Length As Integer
    Dim DummyRecord As RecordScore
    Dim Size As Integer
    
    Length = Len(DummyRecord)
    FName = "\HighScores.rec"
    
    Open App.Path & FName For Random As #1 Len = Length
       Size = LOF(1)
       For X = 1 To Size \ Length
           Get #1, X, Scores(X)
       Next X
    Close #1
    
    NumRecs = Size \ Length
End Sub

Public Sub ShowGrid()
    Dim I As Integer
    Dim J As Integer
    
    BubbleSort Scores(), NumRecs
    
    grdScores.FixedCols = 1
    grdScores.Cols = 3
    grdScores.Rows = MAXENTRIES + 1
    
    For I = 1 To MAXENTRIES
        grdScores.TextMatrix(I, 0) = I
        grdScores.TextMatrix(I, 1) = "Unclaimed"
        grdScores.TextMatrix(I, 2) = VBA.Format$("N/A", "@@@@@@@@@@@@@@@@@@@@@")
    Next
    
    grdScores.TextMatrix(0, 0) = "#:"
    grdScores.TextMatrix(0, 1) = "Player:"
    grdScores.TextMatrix(0, 2) = "Levels Completed:"
    
    grdScores.ColWidth(0) = RANKWIDTH
    
    For I = 0 To grdScores.Rows - 1
        For J = 1 To grdScores.Cols - 1
            grdScores.ColWidth(J) = ((grdScores.Width - RANKWIDTH) / 2) - WIDTHOFFSETX
        Next J
    Next I
    
    For I = 1 To NumRecs
        grdScores.TextMatrix(I, 0) = I
        grdScores.TextMatrix(I, 1) = Scores(I).Name
        grdScores.TextMatrix(I, 2) = VBA.Format$(Scores(I).LevelsComplete, "@@@@@@@@@@@@@@@@@@@@@")
    Next I
End Sub

