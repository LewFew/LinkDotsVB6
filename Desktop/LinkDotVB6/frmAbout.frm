VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LINK DOTS - About"
   ClientHeight    =   5085
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   495
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblExplain 
      Caption         =   $"frmAbout.frx":0000
      Height          =   3015
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   4575
   End
   Begin VB.Label lblDescription 
      Caption         =   $"frmAbout.frx":0376
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   4455
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version 2 (Final)"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   4335
   End
   Begin VB.Label lblAuthor 
      Caption         =   "Author: William Chan"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Author: William Chan
'Date: May 17th, 2019
'Purpose: ICS4U Culminating Assignment

Private Sub Form_Load()

End Sub

Private Sub OKButton_Click()
    Unload Me
End Sub
