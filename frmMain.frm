VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00800000&
   Caption         =   "Notepad 2000"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7950
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   7950
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   3720
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   5895
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   10398
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmMain.frx":08CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox Text2 
      Height          =   1935
      Left            =   2520
      TabIndex        =   1
      Top             =   840
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   3413
      _Version        =   393217
      TextRTF         =   $"frmMain.frx":0992
   End
   Begin MSComctlLib.StatusBar Sbar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   5940
      Width           =   7950
      _ExtentX        =   14023
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6615
            MinWidth        =   6615
            Text            =   "Copyright (c) 2000 Emran Hasan Inc."
            TextSave        =   "Copyright (c) 2000 Emran Hasan Inc."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "1/1/01"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "5:59 PM"
         EndProperty
      EndProperty
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
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu New 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu Open 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu Close 
         Caption         =   "&Close"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu Save 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu SaveAs 
         Caption         =   "Save &As"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu PageSetup 
         Caption         =   "Print Se&tup"
      End
      Begin VB.Menu Print 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu Undo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu Redo 
         Caption         =   "Re&do"
         Shortcut        =   ^Y
      End
      Begin VB.Menu sep4 
         Caption         =   "-"
      End
      Begin VB.Menu Cut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu Copy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu Paste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu sep5 
         Caption         =   "-"
      End
      Begin VB.Menu Find 
         Caption         =   "&Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu FindNext 
         Caption         =   "Find &Next"
         Enabled         =   0   'False
         Shortcut        =   {F3}
      End
      Begin VB.Menu Replace 
         Caption         =   "&Replace"
         Shortcut        =   ^H
      End
      Begin VB.Menu sep20 
         Caption         =   "-"
      End
      Begin VB.Menu Clear 
         Caption         =   "Clea&r"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu SelectAll 
         Caption         =   "&Select All"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuInsert 
      Caption         =   "&Insert"
      Begin VB.Menu Autotext 
         Caption         =   "&Autotext"
         Begin VB.Menu AttLine 
            Caption         =   "Attention Line"
            Begin VB.Menu Attention 
               Caption         =   "Attention:"
            End
            Begin VB.Menu Subject 
               Caption         =   "Subject: Lose of passport"
            End
            Begin VB.Menu ATTN 
               Caption         =   "ATTN:"
            End
         End
         Begin VB.Menu Closing 
            Caption         =   "Closing"
            Begin VB.Menu BestRegards 
               Caption         =   "Best Regards,"
            End
            Begin VB.Menu BestWishes 
               Caption         =   "Best Wishes,"
            End
            Begin VB.Menu Cordially 
               Caption         =   "Cordially,"
            End
            Begin VB.Menu Regards 
               Caption         =   "Regards,"
            End
            Begin VB.Menu WithLove 
               Caption         =   "With Love,"
            End
            Begin VB.Menu SincerelyYours 
               Caption         =   "Sincerely Yours,"
            End
            Begin VB.Menu Sincerely 
               Caption         =   "Sincerely,"
            End
            Begin VB.Menu Thanks 
               Caption         =   "Thanks,"
            End
         End
         Begin VB.Menu MailItem 
            Caption         =   "Mail Item"
            Begin VB.Menu To 
               Caption         =   "To"
            End
            Begin VB.Menu From 
               Caption         =   "From"
            End
            Begin VB.Menu Sub 
               Caption         =   "Sub"
            End
         End
         Begin VB.Menu References 
            Caption         =   "Reference"
            Begin VB.Menu Inreplyto 
               Caption         =   "In reply to:"
            End
            Begin VB.Menu Re 
               Caption         =   "RE:"
            End
            Begin VB.Menu Reference 
               Caption         =   "Reference:"
            End
         End
         Begin VB.Menu Salutation 
            Caption         =   "Salutation"
            Begin VB.Menu DearSir 
               Caption         =   "Dear Sir"
            End
            Begin VB.Menu LadiesandGentlement 
               Caption         =   "Ladies and Gentlement"
            End
            Begin VB.Menu ToWhomItMayConcern 
               Caption         =   "To Whom It May Concern:"
            End
         End
      End
      Begin VB.Menu sep8 
         Caption         =   "-"
      End
      Begin VB.Menu DateTime 
         Caption         =   "&Date and time..."
         Shortcut        =   {F5}
      End
      Begin VB.Menu Symbol 
         Caption         =   "&Symbol"
      End
      Begin VB.Menu sep6 
         Caption         =   "-"
      End
      Begin VB.Menu File 
         Caption         =   "&File"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu Statistics 
         Caption         =   "&Document Statictics"
      End
      Begin VB.Menu sep32 
         Caption         =   "-"
      End
      Begin VB.Menu sEncrypt 
         Caption         =   "&Encryption"
         Shortcut        =   {F6}
      End
      Begin VB.Menu sDecrypt 
         Caption         =   "&Decryption"
         Shortcut        =   {F7}
      End
      Begin VB.Menu sep65 
         Caption         =   "-"
      End
      Begin VB.Menu FullScreen 
         Caption         =   "&Full Screen"
         Shortcut        =   {F12}
      End
      Begin VB.Menu sep54 
         Caption         =   "-"
      End
      Begin VB.Menu Options 
         Caption         =   "&Options"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu Contents 
         Caption         =   "&Contents..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu sep18 
         Caption         =   "-"
      End
      Begin VB.Menu Web 
         Caption         =   "E-mail the author"
      End
      Begin VB.Menu sep27 
         Caption         =   "-"
      End
      Begin VB.Menu About 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const System = 4
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim gblnIgnoreChange As Boolean
Dim gintIndex As Integer
Dim gstrStack(1000) As String
Dim DocChanged As Boolean

Private Sub About_Click()
Form1.Show 1
End Sub

Private Sub Attention_Click()
Text1.SelText = "Attention"
End Sub

Private Sub ATTN_Click()
Text1.SelText = "ATTN:"
End Sub

Private Sub BestRegards_Click()
Text1.SelText = "Best Regards,"
End Sub

Private Sub BestWishes_Click()
Text1.SelText = "Best Wishes,"
End Sub

Private Sub Clear_Click()
Text1.SelText = ""
End Sub

Private Sub Close_Click()
Dim Ans
If DocChanged = True Then
    Ans = MsgBox("Do you want to save the changes ?", vbYesNo + vbQuestion, "Notepad 2000")
        If Ans = vbYes Then
            Call Save_Click
        Else
            Text1.Text = ""
            Text1.Locked = True
        End If
Else
    Text1.Text = ""
    Text1.Locked = True
End If

End Sub

Private Sub Contents_Click()
    htmlhelp hwnd, SetHTMLHelpStrings(), HH_DISPLAY_TOC, 0
    
End Sub

Private Sub Copy_Click()
    'Clears the Clipboard to put text on it
    Clipboard.Clear
    'Sets the Text from rtfText onto the Clipboard
    Clipboard.SetText Text1.SelText
    'Sets the Focus to rtfText
    Text1.SetFocus
End Sub

Private Sub Cordially_Click()
Text1.SelText = "Cordially,"
End Sub

Private Sub Cut_Click()
    'Clears the Clipboard to put text on it
    Clipboard.Clear
    'Sets the Text from rtfText onto the Clipboard
    Clipboard.SetText Text1.SelText
    'Deletes the Selected Text on rtfText
    Text1.SelText = ""
    'Sets the Focus to rtfText
    Text1.SetFocus
    
    DocChanged = True
End Sub

Private Sub DateTime_Click()
Form3.Show 1
End Sub

Private Sub DearSir_Click()
Text1.SelText = "Dear Sir"
End Sub


Private Sub Exit_Click()
Dim Ans
If DocChanged = True Then
    Ans = MsgBox("Do you want to save the changes ?", vbYesNo + vbQuestion, "Notepad 2000")
        If Ans = vbYes Then
            Call Save_Click
            End
        Else
            End
        End If
Else
    End
End If

End Sub

Private Sub File_Click()
Dim sFile As String
Dim txt
On Error Resume Next
With cd1
    .DialogTitle = "Insert file..."
    .CancelError = False
    .Filter = "Text File (*.txt)|*.txt|All Files (*.*)|*.*"
    .ShowOpen
    If Len(.FileName) = 0 Then
        Exit Sub
    End If
    sFile = .FileName
End With
Text2.LoadFile sFile
Text1.SelText = Text2.Text

DocChanged = True
End Sub

Private Sub Find_Click()
    ' Check to see if there are any open documents, if not show error
    If Text1.Text = "" Then
        MsgBox "Document contains no text !", vbInformation, "NotePad 2000"
        Exit Sub
    Else
    frmFind.Show , Me
    End If
End Sub
Private Sub FindNext_Click()

    If Text1.Text = "" Then MsgBox "Document contains no text !", vbInformation, "NotePad 2000": Exit Sub
    On Error GoTo FindNextError
    Dim lngResult As Integer
    Dim lngPos As Integer
    Dim intOptions As Integer

    If frmFind.chkNoHighlight.Value = 1 Then intOptions = intOptions + 8
    If frmFind.chkWholeWord.Value = 1 Then intOptions = intOptions + 2
    If frmFind.chkMatchCase.Value = 1 Then intOptions = intOptions + 4

    lngPos = Text1.SelStart + Text1.SelLength

    lngResult = Text1.Find(frmFind.cboFind.Text, lngPos, , intOptions)

    If lngResult = -1 Then
        MsgBox "Document searched. No such text found !", vbInformation, "NotePad 2000"
        frmFind.cmdFind.Caption = "&Find"
        frmFind.cmdReplace.Enabled = False
        frmFind.cmdReplaceAll.Enabled = False
        FindNext.Enabled = False
    Else
        Text1.SetFocus
    End If
FindNextError:
    Exit Sub
End Sub

Private Sub FullScreen_Click()
Clipboard.SetText Text1.Text
    
    Form5.FSRTB.Text = Clipboard.GetText
    Form5.Show , Me
    Form5.FSRTB.SelStart = 0

    Form5.FSRTB.BackColor = Text1.BackColor
    Form5.FSRTB.Font.Name = Text1.Font.Name
    Form5.FSRTB.Font.Size = Text1.Font.Size
    Form5.FSRTB.SelStart = 0
    Form5.FSRTB.SelLength = Len(Form5.FSRTB.Text)
    Form5.FSRTB.SelColor = Text1.SelColor
    Form5.FSRTB.SelStart = 0
End Sub

Private Sub Options_Click()
Form2.Show 1
End Sub

Private Sub Redo_Click()
    'This is the basic redo stuff.
    gblnIgnoreChange = True
    gintIndex = gintIndex + 1
    On Error Resume Next
    Text1.TextRTF = gstrStack(gintIndex)
    gblnIgnoreChange = False

End Sub

Private Sub Replace_Click()
    
    If Text1.Text = "" Then MsgBox "Document contains no text !", vbInformation, "NotePad 2000": Exit Sub
    With frmFind
        .cmdReplace.Top = 150 'Set cmdReplace top
        .cmdReplace.Caption = "&Replace" 'Set caption
        .lblReplace.Visible = True 'Show lblReplace
        .cboReplace.Visible = True 'Show cboReplace
        .cmdReplaceAll.Visible = True 'Show cmdReplaceAll
        .Show , Me
    End With
    
    DocChanged = True
End Sub

Private Sub Form_Activate()
Text1.SetFocus
End Sub

Private Sub Form_Load()
Dim fontName As String
Dim fontSize As String
Dim Key As String
Dim BlueWhite As String
Dim prg As String
Dim ico As String
prg = App.Path & "\Notepad2000.exe"
ico = App.Path & "\Doc.ico"
Associate prg, "txt", "Text Document", ico
Open App.Path & "\Options.txt" For Input As #1
Input #1, fontName
Input #1, fontSize
Input #1, Key
Input #1, BlueWhite
Close #1

fontName = Right(fontName, (Len(fontName) - 12))
fontSize = Right(fontSize, (Len(fontSize) - 12))

Text1.Font.Name = fontName
Text1.Font.Size = fontSize

If BlueWhite = "0" Then
    Text1.BackColor = vbWhite
    Text1.SelColor = vbBlack
ElseIf BlueWhite = "1" Then
    Text1.BackColor = &H800000
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
    Text1.SelColor = vbWhite
End If

Me.Caption = "Notepad 2000 - Untitled"

If Command = "" Then
    Exit Sub
Else
    Text1.LoadFile Command
End If

DocChanged = False
End Sub

Private Sub Form_Resize()
On Error Resume Next
Text1.Height = Me.Height - 975
Text1.Width = Me.Width - 135
Sbar1.Panels.Item(1).MinWidth = Me.Width - 4320
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub From_Click()
Text1.SelText = "From"
End Sub

Private Sub Inreplyto_Click()
Text1.SelText = "In reply to:"
End Sub

Private Sub LadiesandGentlement_Click()
Text1.SelText = "Ladies and Gentlement"
End Sub

Private Sub New_Click()
Dim Ans
If DocChanged = True Then
    Ans = MsgBox("Do you want to save the changes ?", vbYesNo + vbQuestion, "Notepad 2000")
        If Ans = vbYes Then
            Call Save_Click
            Text1.Text = ""
            DocChanged = False
        Else
            Text1.Text = ""
            DocChanged = False
        End If
Else
    Text1.Text = ""
    DocChanged = False
End If
End Sub

Private Sub Open_Click()
Dim Ans
If DocChanged = True Then
    Ans = MsgBox("Do you want to save the changes ?", vbYesNo + vbQuestion, "Notepad 2000")
        If Ans = vbYes Then
            Call Save_Click
            Text1.Text = ""
            Call OpenFile
            Text1.Locked = False
            DocChanged = False
        Else
            Text1.Text = ""
            Call OpenFile
            Text1.Locked = False
            DocChanged = False
        End If
Else
    Text1.Text = ""
    Call OpenFile
    Text1.Locked = False
    DocChanged = False
End If

End Sub

Private Sub PageSetup_Click()
On Error Resume Next

    cd1.ShowPrinter  'Show Page Setup dialog
End Sub

Private Sub Paste_Click()
    'Puts the Text from the clipboard into rtfText
    Text1.SelText = Clipboard.GetText
    'Sets the Focus to rtfText
    Text1.SetFocus
    If Text1.BackColor = vbWhite Then
        Text1.SelStart = 0
        Text1.SelLength = Len(Text1.Text)
        Text1.SelColor = vbBlack
        Text1.SelStart = 0
    Else
        Text1.SelStart = 0
        Text1.SelLength = Len(Text1.Text)
        Text1.SelColor = vbWhite
        Text1.SelStart = 0
    End If
    DocChanged = True
End Sub

Private Sub Print_Click()
On Error GoTo PrintCanceled:

  cd1.Flags = DefaultFlags
  cd1.ShowPrinter
  Printer.Copies = cd1.Copies

  If Text1.SelText = "" Then
    Printer.Print Text1.Text
  Else
    Printer.Print Text1.SelText
  End If

PrintCanceled:

End Sub

Private Sub RE_Click()
Text1.SelText = "RE:"
End Sub

Private Sub Reference_Click()
Text1.SelText = "Reference:"
End Sub

Private Sub Regards_Click()
Text1.SelText = "Regards,"
End Sub
Private Sub Save_Click()
Call SaveFile
End Sub

Private Sub SaveAs_Click()
Call SaveFile
End Sub

Private Sub sDecrypt_Click()
Dim KeyString As String
Dim a, b, d
Open App.Path & "\Options.txt" For Input As #1
Input #1, a
Input #1, b
Input #1, KeyString
Input #1, d
Close #1

Text1.Text = modCrypt.DecryptText(Text1.Text, KeyString)
End Sub

Private Sub SelectAll_Click()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
Text1.SetFocus
End Sub

Private Sub sEncrypt_Click()
Dim KeyString As String
Dim a, b, d

Open App.Path & "\Options.txt" For Input As #1
Input #1, a
Input #1, b
Input #1, KeyString
Input #1, d
Close #1

Text1.Text = modCrypt.EncryptText(Text1.Text, KeyString)
End Sub

Private Sub Sincerely_Click()
Text1.SelText = "Sincerely,"
End Sub

Private Sub SincerelyYours_Click()
Text1.SelText = "Sincerely Yours,"
End Sub


Private Sub Statistics_Click()
frmStatistics.Show 1
End Sub

Private Sub Sub_Click()
Text1.SelText = "Sub"
End Sub

Private Sub Subject_Click()
Text1.SelText = "Seubject: Lost ID"
End Sub

Private Sub Symbol_Click()
Form4.Show , Me
End Sub

Private Sub Text1_Change()
    'Basically this updates the Undo and Redo
    If Not gblnIgnoreChange Then
        gintIndex = gintIndex + 1
        gstrStack(gintIndex) = Text1.TextRTF
    End If

Text2.Text = Text1.Text
DocChanged = True
End Sub

Private Sub Thanks_Click()
Text1.SelText = "Thanks,"
End Sub

Private Sub To_Click()
Text1.SelText = "To"
End Sub

Private Sub ToWhomItMayConcern_Click()
Text1.SelText = "To Whom it may concern:"
End Sub

Private Sub Undo_Click()
    'This says that if the Index is = to 0, then It shouldn't undo anymore
    If gintIndex = 0 Then Exit Sub
    
    'This is the basic undo stuff.
    gblnIgnoreChange = True
    gintIndex = gintIndex - 1
    On Error Resume Next
    Text1.TextRTF = gstrStack(gintIndex)
    gblnIgnoreChange = False

End Sub

Private Sub SaveFile()
Dim sFile As String
On Error Resume Next
With cd1
    .DialogTitle = "Save As..."
    .CancelError = False
    .Filter = "Text File (*.txt)|*.txt|All Files (*.*)|*.*"
    .ShowSave
    If Len(.FileName) = 0 Then
        Exit Sub
    End If
    sFile = .FileName
End With
Text1.SaveFile sFile
End Sub
Private Sub OpenFile()
Dim sFile As String
On Error Resume Next
With cd1
    .DialogTitle = "Open..."
    .CancelError = False
    .Filter = "Text File (*.txt)|*.txt|All Files (*.*)|*.*"
    .ShowOpen
    If Len(.FileName) = 0 Then
        Exit Sub
    End If
    sFile = .FileName
End With
Text1.LoadFile sFile
Me.Caption = "Notepad 2000 - " & cd1.FileTitle
If Text1.BackColor = vbWhite Then
    Call SelectAll_Click
    Text1.SelColor = vbBlack
    Text1.SelStart = 0
Else
    Call SelectAll_Click
    Text1.SelColor = vbWhite
    Text1.SelStart = 0
End If
End Sub

Private Sub Web_Click()
    Call ShellExecute(&O0, vbNullString, "mailto:webmaster@emranhasan.com", vbNullString, vbNullString, vbNormalFocus)
End Sub

Private Sub WithLove_Click()
Text1.SelText = "With Love,"
End Sub

