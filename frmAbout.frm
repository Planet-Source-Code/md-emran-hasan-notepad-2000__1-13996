VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About..."
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4350
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "http://www.emranhasan.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   960
      MouseIcon       =   "frmAbout.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   7
      ToolTipText     =   "Vsit author's home page on the net"
      Top             =   3000
      Width           =   2085
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "webmaster@emranhasan.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   960
      MouseIcon       =   "frmAbout.frx":0316
      MousePointer    =   99  'Custom
      TabIndex        =   6
      ToolTipText     =   "E-mail the author"
      Top             =   2640
      Width           =   2145
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   960
      TabIndex        =   4
      Top             =   3480
      Width           =   555
   End
   Begin VB.Label Label4 
      Caption         =   $"frmAbout.frx":0620
      Height          =   1095
      Left            =   960
      TabIndex        =   3
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "Copyright (c) 2000 Md Emran Hasan."
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Version 1.0.12"
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   575
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Notepad 2000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   275
      Width           =   1455
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Picture         =   "frmAbout.frx":06DD
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetTickCount Lib "kernel32.dll" () As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim lngTickCount As Long
Dim HTime As Long

lngTickCount = GetTickCount
HTime = lngTickCount / 1000 / 60

Label5.Caption = "Windows running for " & Left(HTime, 5) & " minutes."
End Sub


Private Sub Label6_Click()
    Call ShellExecute(&O0, vbNullString, "mailto:webmaster@emranhasan.com", vbNullString, vbNullString, vbNormalFocus)
End Sub

Private Sub Label7_Click()
    Call ShellExecute(&O0, vbNullString, "http://www.emranhasan.com", vbNullString, vbNullString, vbNormalFocus)
End Sub
