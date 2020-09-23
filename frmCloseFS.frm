VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form6 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Close Full Screen"
   ClientHeight    =   360
   ClientLeft      =   10470
   ClientTop       =   1305
   ClientWidth     =   510
   Icon            =   "frmCloseFS.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   360
   ScaleWidth      =   510
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList imgFull 
      Left            =   2400
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCloseFS.frx":000C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrFull 
      Height          =   330
      Left            =   75
      TabIndex        =   0
      Top             =   0
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "imgFull"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "FullScreen"
            Object.ToolTipText     =   "Close Full Screen"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
    MakeTopMost Me.hWnd
    DisableX Me
    Form5.FSRTB.SelStart = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MakeNormal Me.hWnd
End Sub

Private Sub tbrFull_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "FullScreen"
            Unload Form5
            Unload Me
            frmMain.SetFocus
    End Select
End Sub

