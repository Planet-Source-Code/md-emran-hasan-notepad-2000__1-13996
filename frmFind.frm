VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   2640
   ClientLeft      =   3390
   ClientTop       =   5805
   ClientWidth     =   5625
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboFind 
      Height          =   315
      Left            =   1245
      TabIndex        =   11
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Default         =   -1  'True
      Height          =   375
      Left            =   4395
      TabIndex        =   10
      Top             =   120
      Width           =   1110
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4395
      TabIndex        =   9
      Top             =   600
      Width           =   1110
   End
   Begin VB.ComboBox cboReplace 
      Height          =   315
      Left            =   1245
      TabIndex        =   8
      Top             =   600
      Width           =   3015
   End
   Begin VB.PictureBox picBar 
      BorderStyle     =   0  'None
      Height          =   1530
      Left            =   120
      ScaleHeight     =   1530
      ScaleWidth      =   5460
      TabIndex        =   0
      Top             =   1080
      Width           =   5460
      Begin VB.Frame Frame1 
         Caption         =   "Search Options"
         Height          =   1455
         Left            =   75
         TabIndex        =   4
         Top             =   0
         Width           =   4065
         Begin VB.CheckBox chkNoHighlight 
            Caption         =   "No &Highlight"
            Height          =   240
            Left            =   150
            TabIndex        =   7
            Top             =   1080
            Width           =   1965
         End
         Begin VB.CheckBox chkMatchCase 
            Caption         =   "Match Ca&se"
            Height          =   240
            Left            =   150
            TabIndex        =   6
            Top             =   685
            Width           =   1965
         End
         Begin VB.CheckBox chkWholeWord 
            Caption         =   "Find Whole Word &Only"
            Height          =   240
            Left            =   150
            TabIndex        =   5
            Top             =   300
            Width           =   1965
         End
      End
      Begin VB.CommandButton cmdReplace 
         Caption         =   "&Replace..."
         Height          =   375
         Left            =   4275
         TabIndex        =   3
         Top             =   150
         Width           =   1110
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "&Help"
         Height          =   375
         Left            =   4275
         TabIndex        =   2
         Top             =   1080
         Width           =   1110
      End
      Begin VB.CommandButton cmdReplaceAll 
         Caption         =   "Replace &All"
         Height          =   375
         Left            =   4275
         TabIndex        =   1
         Top             =   600
         Width           =   1110
      End
   End
   Begin VB.Label lblFind 
      Caption         =   "Fin&d What:"
      Height          =   240
      Left            =   120
      TabIndex        =   13
      Top             =   195
      Width           =   840
   End
   Begin VB.Label lblReplace 
      Caption         =   "Replace &With:"
      Height          =   240
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   1065
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdFind_Click()
    On Error GoTo FindError
    Dim lngResult As Long
    Dim lngPos As Long
    Dim intOptions As Integer
    ' Set search options
    If chkNoHighlight.Value = 1 Then intOptions = intOptions + 8
    If chkWholeWord.Value = 1 Then intOptions = intOptions + 2
    If chkMatchCase.Value = 1 Then intOptions = intOptions + 4

    If cmdFind.Caption = "&Find" Then 'If first time
        ' Get position of the searched word
        lngResult = frmMain.Text1.Find(cboFind.Text, 0, , intOptions)

        If lngResult = -1 Then 'Text not found
        MsgBox "Document searched. No such text found !", vbInformation, "NotePad 2000"
            cmdFind.Caption = "&Find" 'Set caption
            frmMain.FindNext.Enabled = False 'Disable Find Next menu
        Else 'Text found
            frmMain.Text1.SetFocus 'Set focus to rtfText
            cmdReplace.Enabled = True 'Enable Replace button
            cmdReplaceAll.Enabled = True 'Enable ReplaceAll button
            cmdFind.Caption = "&Find Next" 'Set caption
            frmMain.FindNext.Enabled = True 'Enable Find Next menu
        End If
    Else 'Find Next
        lngPos = frmMain.Text1.SelStart + frmMain.Text1.SelLength
        lngResult = frmMain.Text1.Find(cboFind.Text, lngPos, , intOptions)

        If lngResult = -1 Then 'Text not found
         MsgBox "Document searched. No such text found !", vbInformation, "NotePad 2000"
            cmdFind.Caption = "&Find" 'Set caption
            cmdReplace.Enabled = False 'Disable Replace button
            cmdReplaceAll.Enabled = False 'Disable ReplaceAll button
            frmMain.FindNext.Enabled = False 'Disable Find Next menu
        Else 'Text found
            frmMain.Text1.SetFocus 'Set focus to rtfText
            frmMain.FindNext.Enabled = True 'Enable Find Next menu
        End If
    End If
FindError:
    Exit Sub
End Sub

Private Sub cmdReplace_Click()
    On Error GoTo ReplaceError
    Dim lngResult As Long
    Dim lngPos As Long
    Dim intOptions As Integer
    
    If cmdReplace.Caption = "&Replace..." Then 'Show replace
        cmdReplace.Top = 150 'Set cmdReplace top
        cmdReplace.Caption = "&Replace" 'Set caption
        lblReplace.Visible = True 'Show lblReplace
        cboReplace.Visible = True 'Show cboReplace
        cmdReplaceAll.Visible = True 'Show cmdReplaceAll
        Exit Sub
    End If

    ' Set search options
    If chkNoHighlight.Value = 1 Then intOptions = intOptions + 8
    If chkWholeWord.Value = 1 Then intOptions = intOptions + 2
    If chkMatchCase.Value = 1 Then intOptions = intOptions + 4
    
    With frmMain
        .Text1.SelText = cboReplace.Text 'Replace text
        ' Find next
        lngPos = .Text1.SelStart + .Text1.SelLength
        ' Get position of the searched word
        lngResult = .Text1.Find(cboFind.Text, lngPos, , intOptions)

        If lngResult = -1 Then 'Text not found
        MsgBox "Document searched. No such text found !", vbInformation, "NotePad 2000"
            cmdFind.Caption = "&Find" 'Set caption
            cmdReplace.Enabled = False 'Disable Replace button
            cmdReplaceAll.Enabled = False 'Disable ReplaceAll button
        Else 'Text found
            .Text1.SetFocus 'Set focus
        End If
    End With
ReplaceError:
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdReplaceAll_Click()
    On Error GoTo ReplaceAllError
    Dim intCount As Integer
    Dim lngPos As Long
    Dim intOptions As Integer
    ' Set search options
    If chkNoHighlight.Value = 1 Then intOptions = intOptions + 8
    If chkWholeWord.Value = 1 Then intOptions = intOptions + 2
    If chkMatchCase.Value = 1 Then intOptions = intOptions + 4

    intCount = 0
    lngPos = 0
    With frmMain
        Do
            If .Text1.Find(cboFind.Text, lngPos, , intOptions) = -1 Then 'Text not fount
                If intCount > 0 Then 'Show how many replacments have been made
                    MsgBox "The specified region has been searched. " & vbCrLf & intCount & " replacements have been made.", vbInformation, "Notepad 2000"
                End If
                cmdFind.Caption = "&Find" 'Set caption
                cmdReplace.Enabled = False 'Disable Replace button
                cmdReplaceAll.Enabled = False 'Disable ReplaceAll button
                Exit Do
            Else 'Text found
                lngPos = .Text1.SelStart + .Text1.SelLength
                intCount = intCount + 1 'Increase counter by 1
                .Text1.SelText = cboReplace.Text 'Replace text
            End If
        Loop
    End With
ReplaceAllError:
        Exit Sub
End Sub

Private Sub Form_Load()
    cmdReplace.Top = 525 'Set cmdReplace top
    lblReplace.Visible = False 'Hide lblReplace
    cboReplace.Visible = False 'Hide cboReplace
    cmdReplaceAll.Visible = False 'Hide cmdReplaceAll
    
    cboFind.AddItem frmMain.Text1.SelText 'Add selected text to find combobox
    cboFind.Text = frmMain.Text1.SelText 'Set text in cbo
End Sub

