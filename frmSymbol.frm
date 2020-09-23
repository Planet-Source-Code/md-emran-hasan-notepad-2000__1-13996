VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Insert Symbol"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2985
   LinkTopic       =   "Form4"
   ScaleHeight     =   3990
   ScaleWidth      =   2985
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Symbols"
      Height          =   3375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2895
      Begin VB.ListBox lstSymbols 
         Height          =   2985
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "&Insert"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3480
      Width           =   1215
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub lstSymbols_DblClick()
    frmMain.Text1.SelText = Right(lstSymbols.Text, 1)
End Sub

Private Sub cmdInsert_Click()
frmMain.Text1.SelText = Right(lstSymbols.Text, 1)
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim I As Integer
    
    'Set font name
    lstSymbols.FontName = frmMain.Text1.Font.Name
    For I = 1 To 255
        ' Fills lstSymbols with Symbols
        If I < 10 Then
            lstSymbols.AddItem I & "     -  " & Chr(I)
        ElseIf I < 100 Then
            lstSymbols.AddItem I & "   -  " & Chr(I)
        Else
            lstSymbols.AddItem I & " -  " & Chr(I)
        End If
    Next I
End Sub
