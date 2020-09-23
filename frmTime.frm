VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Date/Time"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4200
   Icon            =   "frmTime.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   4200
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Formats"
      Height          =   2775
      Left            =   10
      TabIndex        =   2
      Top             =   10
      Width           =   2775
      Begin VB.ListBox Lsttime 
         Height          =   2400
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmMain.Text1.SelText = Lsttime.Text
Unload Me
End Sub

Private Sub Command2_Click()
Me.Hide
End Sub

Private Sub Form_Load()
Lsttime.AddItem Format(Now, "long time")
Lsttime.AddItem Format(Now, "short time")
Lsttime.AddItem Format(Now, "medium time")
Lsttime.AddItem Format(Now, "general date")
Lsttime.AddItem Format(Now, "long date")
Lsttime.AddItem Format(Now, "medium date")
Lsttime.AddItem Format(Now, "short date")
Lsttime.AddItem (Date)
Lsttime.AddItem Format(Date, "dd - mm - yyyy")
Lsttime.AddItem Format(Date, "dd-mm-yy")
Lsttime.AddItem Format(Date, "dd/mm/yy")
Lsttime.AddItem Format(Date, "dd/mm/yyyy")
Lsttime.AddItem Format(Date, "dd/mm")
Lsttime.AddItem Format(Date, "dd")
Lsttime.AddItem Format(Time, "hh-mm-ss")
Lsttime.AddItem Format(Time, "hh.mm.ss")
Lsttime.AddItem Format(Time, "hh-mm")
End Sub

Private Sub Lsttime_DblClick()
Command1_Click
End Sub

