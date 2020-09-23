VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "About Search"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   735
      Left            =   2520
      ScaleHeight     =   675
      ScaleWidth      =   795
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton SaveCommand 
      Caption         =   "Save"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   2280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton PrintCommand 
      Caption         =   "Print"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   1680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   10185
      Left            =   480
      Picture         =   "frmAbout.frx":0442
      ScaleHeight     =   10125
      ScaleWidth      =   13500
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   13560
   End
   Begin VB.Label Pietro_Cecchi 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pietro Cecchi  cecchi@infinito.it  pietrocecchi@inwind.it"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   2055
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this form I added for the ones who like pictures:
'how to size, view, print and save an image

'used methods: PaintPicture, SavePicture

Private Sub Form_Load()
 'leave this sub empty, all is done in Form_Resize
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Help "frmAbout"
End Sub

Private Sub Form_Resize()
 Static nexttime As Boolean
 
 If Not nexttime Then 'firsttime
   With Picture1
     .Visible = False
     .AutoSize = True 'be sure is set true also at project time
     .AutoRedraw = True
   End With
   With Picture2
     .Visible = False
     .AutoSize = False
     .AutoRedraw = True
   End With
   nexttime = True
End If
 
 PrintCommand.Visible = False
 SaveCommand.Visible = False
 Pietro_Cecchi.Visible = False
 Picture = LoadPicture()
 DoEvents
 
 'put PrintCommand in place
 With PrintCommand
   .Move ScaleWidth - .Width - ScaleHeight / 10, ScaleHeight - .Height - ScaleHeight / 10
 End With
 'put SaveCommand in place
 With SaveCommand
   .Move ScaleWidth - .Width * 2 - ScaleHeight / 10 * 2, ScaleHeight - .Height - ScaleHeight / 10
 End With
 'resize picture
 With Picture1 'the picture you will see is frmAbout.Picture
   PaintPicture .Picture, 0, 0, ScaleWidth, ScaleHeight, 0, 0, .Width, .Height
   Picture = Image
 End With
 
 PrintCommand.Visible = True
 SaveCommand.Visible = True
 Pietro_Cecchi.Visible = True

End Sub

Private Sub Pietro_Cecchi_Click()
  Shell "start mailto:pietrocecchi@inwind.it"
End Sub

Private Sub Pietro_Cecchi_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Help "Pietro_Cecchi"
End Sub

Private Sub PrintCommand_Click() 'print
 PrintCommand.Visible = False
 On Error Resume Next
 
  If Width >= Height Then
   'landscape
   Printer.Orientation = vbPRORLandscape
  Else
   'portrait
   Printer.Orientation = vbPRORPortrait
  End If
  delta = Screen.TwipsPerPixelX '1 pixel width
  'Draw a box around whole printer area
  Printer.Line (delta, delta)-(Printer.ScaleWidth - delta, Printer.ScaleHeight - delta), RGB(0, 0, 255), B  'Draw box
  'paint image in center of printer area, take dimensions of about form
  X = (Printer.ScaleWidth - ScaleWidth) / 2
  Y = (Printer.ScaleHeight - ScaleHeight) / 2
  With Picture1
   Printer.PaintPicture .Picture, X, Y, ScaleWidth, ScaleHeight, 0, 0, .Width, .Height
  End With
  'issue print command
  Printer.EndDoc
 
 Err.Clear
 On Error GoTo 0
 PrintCommand.Visible = True
End Sub

Private Sub PrintCommand_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Help "PrintCommand"
End Sub

Private Sub SaveCommand_Click()
 
 If SaveCommand.Caption = "Done" Then Exit Sub
 SaveCommand.Caption = "Wait..."
 
 With Picture2 'the picture you will save is Picture2
   .Width = ScaleWidth: Picture2.Height = ScaleHeight
   .PaintPicture Picture1.Picture, 0, 0, ScaleWidth, ScaleHeight, 0, 0, Picture1.Width, Picture1.Height
   .Picture = .Image
 End With

 SavePicture Picture2, "c:\Search About.bmp"
 
 SaveCommand.Caption = "Done"
  Pause 2000
 SaveCommand.Caption = "Save"
 
End Sub

Private Sub SaveCommand_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Help "SaveCommand"
End Sub
