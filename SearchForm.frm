VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form SearchForm 
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Files and/or Dirs of given partial name, into given Dir"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   6960
   Icon            =   "SearchForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   6960
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   150
      Left            =   0
      TabIndex        =   1
      Top             =   -120
      Width           =   6975
   End
   Begin MSComctlLib.StatusBar HelpWindow 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   6015
      Width           =   6960
      _ExtentX        =   12277
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12224
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   5655
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   6735
      Begin VB.ListBox FindFilesTmpResults 
         Height          =   450
         Left            =   3840
         TabIndex        =   9
         Top             =   2760
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox nametosearch 
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   4560
         TabIndex        =   8
         Text            =   "nametosearch"
         Top             =   600
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Go search!"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1080
         Width           =   2055
      End
      Begin VB.DriveListBox Drive1 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   4215
      End
      Begin VB.DirListBox Dir1 
         ForeColor       =   &H00C00000&
         Height          =   1215
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   4215
      End
      Begin VB.ListBox FindFilesTmpDirs 
         Height          =   450
         Left            =   3840
         TabIndex        =   4
         Top             =   3600
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Stop search!!!"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1800
         Width           =   1815
      End
      Begin VB.ListBox FindFilesResults 
         ForeColor       =   &H000000C0&
         Height          =   2985
         Left            =   120
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   2520
         Width           =   6495
      End
      Begin VB.Label Drive1Label 
         BackStyle       =   0  'Transparent
         Height          =   615
         Left            =   0
         TabIndex        =   11
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label Label1 
         Caption         =   "Search into dir:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Name to search:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   4560
         TabIndex        =   14
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "Results:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label SearchInterruptedLabel 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Caption         =   "Search interrupted by program"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2160
         TabIndex        =   12
         Top             =   2280
         Width           =   2175
      End
   End
   Begin VB.Menu menufile 
      Caption         =   "File"
      Begin VB.Menu menuexit 
         Caption         =   "Exit"
      End
      Begin VB.Menu menuline 
         Caption         =   "-"
      End
      Begin VB.Menu menucancel 
         Caption         =   "Cancel"
      End
   End
   Begin VB.Menu menuabout 
      Caption         =   "About"
   End
   Begin VB.Menu menupopmenu 
      Caption         =   "popmenu"
      Visible         =   0   'False
      Begin VB.Menu menupopmenucopy 
         Caption         =   "copy selected item(s) to clipboard"
      End
      Begin VB.Menu menupopmenuline 
         Caption         =   "-"
      End
      Begin VB.Menu menupopmenucancel 
         Caption         =   "cancel"
      End
   End
End
Attribute VB_Name = "SearchForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   FindFilesResults.BackColor = vbYellow
   FindFilesTmpResults.Clear
   FindFilesTmpDirs.Clear
   FindFilesResults.Clear
   SearchFilesInDir Dir1.Path, nametosearch
   FindFilesResults.BackColor = vbWindowBackground
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Help "Command1"
End Sub

Private Sub Command2_Click()
   SearchInterrupted = True
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Help "Command2"
End Sub

Private Sub Dir1_Click()
   With Dir1
      .Path = .List(.ListIndex)
   End With
End Sub

Private Sub Dir1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Help "Dir1"
End Sub

Private Sub Drive1_Change()
   Dir1.Path = Left(Drive1.Drive, 2) & "\"
End Sub

Private Sub Drive1Label_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Help "Drive1Label"
End Sub

Private Sub FindFilesResults_DblClick()
  SendKeys Chr(123) & "home" & Chr(125) & "+" & Chr(123) & "end" & Chr(125)
End Sub

Private Sub FindFilesResults_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button <> 2 Then Exit Sub
 With FindFilesResults
   Select Case .SelCount
      Case Is <= 0
         menupopmenucopy.Caption = "Copy selected item(s) (up to 3000) to clipboard"
      Case 1
         menupopmenucopy.Caption = "Copy selected item to clipboard"
      Case Else
         menupopmenucopy.Caption = "Copy selected items (up to 3000, they are " & .SelCount & ") to clipboard"
   End Select
   menupopmenucopy.Enabled = (.SelCount > 0) And (.SelCount <= 3000)
 End With
   PopupMenu menupopmenu
End Sub

Private Sub FindFilesResults_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim MouseIsOn
  If Button <> 0 Then Exit Sub
  With FindFilesResults
     Set Font = .Font
     MouseIsOn = Int(.TopIndex + Y / TextHeight("Hello World!") + 1)
     If MouseIsOn > .ListCount Then MouseIsOn = ""
  End With
    Help "FindFilesResults", MouseIsOn
End Sub

Private Sub Form_Load()
   'adds horizontal scroll bar to list box
   nNewWidth = FindFilesResults.Width + 100 'new width in pixels
   nRet = SendMessage(FindFilesResults.hwnd, LB_SETHORIZONTALEXTENT, nNewWidth, ByVal 0&)
   
   'assure a finite number of items will be displayed
   'do it at project time
   'FindFilesResults.IntegralHeight = True
   
   'initial set of dir1.path
   Drive1_Change
   'initial search name (wildcars allowed)
   nametosearch.Text = "*.*"

   SearchInterruptedLabel.Visible = False
   FindFilesTmpResults.Visible = False
   FindFilesTmpDirs.Visible = False
   
   'special note:
   ' because drive1 doesn't support the
   ' mousemove event, Drive1Label is used.
   
End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Help "SearchForm"
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Help "Frame2"
End Sub

Private Sub HelpWindow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Help "HelpWindow"
End Sub

Private Sub menuabout_Click()
  frmAbout.Show
End Sub

Private Sub menuexit_Click()
  End
End Sub

Private Sub menupopmenucopy_Click()
  'copy selected items to clipboard
  Clipboard.Clear
  With FindFilesResults
     ncount = 0
     SearchInterruptedLabel.Caption = "Clipboard 0%"
     SearchInterruptedLabel.Visible = True
     SearchForm.MousePointer = vbHourglass
     DoEvents
     SearchForm.Refresh
     Frame2.Enabled = False
     For a = 0 To .ListCount - 1
       DoEvents
        If .Selected(a) Then
           ncount = ncount + 1
           If ncount = 1 Then
              msg = .List(a)
           Else
              msg = vbNewLine & .List(a)
           End If
           Clipboard.SetText Clipboard.GetText & msg
        End If
      SearchInterruptedLabel.Caption = "Clipboard " & Format(a / (.ListCount - 1), "##.#%")
     Next
     SearchInterruptedLabel.Visible = False
     Frame2.Enabled = True
     SearchForm.MousePointer = vbDefault
  
  End With
'  Clipboard.Clear
'  Clipboard.SetText msg

End Sub

Private Sub nametosearch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Help "nametosearch"
End Sub
