VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3705
   Icon            =   "frmOption.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   3705
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picEnc 
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   240
      ScaleHeight     =   1695
      ScaleWidth      =   3165
      TabIndex        =   9
      Top             =   480
      Visible         =   0   'False
      Width           =   3165
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   0
         MaxLength       =   10
         TabIndex        =   11
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label4 
         Caption         =   "* Maximum 10 character."
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "&Encryption Key"
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   2175
      End
   End
   Begin VB.PictureBox picEditor 
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   240
      ScaleHeight     =   1815
      ScaleWidth      =   3255
      TabIndex        =   3
      Top             =   480
      Width           =   3255
      Begin VB.CheckBox Check1 
         Caption         =   "&Blue background, white text."
         Height          =   255
         Left            =   0
         TabIndex        =   8
         Top             =   1440
         Width           =   3015
      End
      Begin VB.ComboBox cboSIze 
         Height          =   315
         Left            =   0
         TabIndex        =   7
         Top             =   960
         Width           =   3135
      End
      Begin VB.ComboBox cboFont 
         Height          =   315
         Left            =   0
         TabIndex        =   5
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label2 
         Caption         =   "Font &Size"
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "&Font"
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin MSComctlLib.TabStrip Tab1 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   4048
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Editor"
            Key             =   "Editor"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Encryption"
            Key             =   "Encryption"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim fontName As String
Dim fontSize As String
Dim Key As String
Dim BlueWhite
Open App.Path & "\Options.txt" For Output As #1
fontName = "Font Name : " & cboFont.Text
fontSize = "Font Size : " & cboSIze.Text
Key = Text1.Text
BlueWhite = Check1.Value
Print #1, fontName
Print #1, fontSize
Print #1, Key
Print #1, BlueWhite
Close #1
Unload Me
frmMain.Text1.Font.Name = cboFont.Text
frmMain.Text1.Font.Size = cboSIze.Text

If BlueWhite = "0" Then
    frmMain.Text1.BackColor = vbWhite
    frmMain.Text1.SelColor = vbBlack
ElseIf BlueWhite = "1" Then
    frmMain.Text1.BackColor = &H800000
    frmMain.Text1.SelStart = 0
    frmMain.Text1.SelLength = Len(frmMain.Text1.Text)
    frmMain.Text1.SelColor = vbWhite
End If

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim fontName As String
Dim fontSize As String
Dim Key As String
Dim BlueWhite As String
Dim a As String
Dim b As String
Dim c As String
Dim d As String
Open App.Path & "\Options.txt" For Input As #1
Input #1, a
Input #1, b
Input #1, c
Input #1, d
Close #1
        
    For x = 1 To Screen.FontCount
        Form2.cboFont.AddItem Screen.Fonts(x)
    Next
    
    For i = 5 To 72
        ' Fills combobox with font size from 5 to 75
        Form2.cboSIze.AddItem i
    Next i
    

fontName = Right(a, (Len(a) - 12))
fontSize = Right(b, (Len(b) - 12))
Text1.Text = c
cboFont.Text = fontName
cboSIze.Text = fontSize
Check1.Value = d
End Sub

Private Sub Tab1_Click()
If Tab1.SelectedItem.Caption = "Editor" Then
    picEditor.Visible = True
    picEnc.Visible = False
ElseIf Tab1.SelectedItem.Caption = "Encryption" Then
    picEditor.Visible = False
    picEnc.Visible = True
End If
End Sub
