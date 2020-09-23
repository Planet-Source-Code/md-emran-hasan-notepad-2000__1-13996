VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form5 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Form5"
   ClientHeight    =   5805
   ClientLeft      =   3150
   ClientTop       =   1170
   ClientWidth     =   7215
   LinkTopic       =   "Form5"
   ScaleHeight     =   5805
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   Begin RichTextLib.RichTextBox FSRTB 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   5318
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"frmFScreen.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Form_Resize
    FSRTB.SelStart = 0
    Form6.Show , Me
End Sub
Private Sub Form_Resize()
    Me.Left = -30
    Me.Top = -30
    Me.Height = Screen.Height + 30
    Me.Width = Screen.Width + 30
    FSRTB.Width = Me.Width - 30
    FSRTB.Height = Me.Height - 30
End Sub
Private Sub Form_Unload(Cancel As Integer)
        
        Clipboard.SetText FSRTB.Text
        frmMain.Text1.Text = Clipboard.GetText
        
        frmMain.Text1.SelStart = 0
        FSRTB.Text = ""
        Exit Sub
    
End Sub
