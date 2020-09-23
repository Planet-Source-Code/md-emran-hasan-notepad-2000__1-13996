VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Replace"
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5250
   Icon            =   "frmReplace.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   5250
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Replace"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Replace &With"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Find What:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   780
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call frmMain.Text1.ReplaceText(Text1.Text, Text2.Text)
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Me.Text1.SetFocus
End Sub
Private Sub Label1_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Text1.SetFocus
End Sub

Private Sub Label2_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Text2.SetFocus
End Sub

