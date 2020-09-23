VERSION 5.00
Begin VB.Form frmStatistics 
   Caption         =   "Document Statistics"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3210
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   3210
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   25
      Left            =   1560
      TabIndex        =   10
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   920
      TabIndex        =   0
      Top             =   2600
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Document Statistics"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1410
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "File Name"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   705
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Total Words"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   870
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Total Character"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Total Lines"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   780
   End
   Begin VB.Label lblLine 
      Caption         =   " "
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Char 
      Caption         =   " "
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Word 
      Caption         =   " "
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label FileName 
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
End
Attribute VB_Name = "frmStatistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim I As Integer
 Dim ReadLine As Integer
  For I = 1 To Len(frmMain.Text1.Text)
   ch = Mid(frmMain.Text1.Text, I, 1)
    If ch = Chr(10) Then
     ReadLine = ReadLine + 1
     End If
     Next
lblLine.Caption = ReadLine + 1
Char.Caption = Len(frmMain.Text1.Text)
Call TotalWords
FileName.Caption = Mid(frmMain.Caption, 16)
End Sub

Private Sub TotalWords()
Dim Position As Long
Dim words As Long
Dim myText As String
    Position = 1
    myText = frmMain.Text1.Text
    myText = Replace(myText, Chr(13) & Chr(10), " ")
    myText = Replace(myText, Chr(9), " ")
    myText = Trim(myText)
    If Len(myText) > 0 Then words = 1
    Do While Position > 0
        Position = InStr(Position, myText, " ")
        If Position > 0 Then
            words = words + 1
            While Mid(myText, Position, 1) = " "
                Position = Position + 1
            Wend
        End If
    Loop
    Word.Caption = words
End Sub

