Attribute VB_Name = "HelpSub"
Public Sub Help(ByRef SourceObject As String, Optional ByRef morehelp = "")
'processes all help messages, basing on mousemove events

With SearchForm

Select Case SourceObject
  Case "HelpWindow" 'help window
     msg = "This is the help window..."
  Case "SearchForm", "Frame2" 'form and frame
     msg = ""
  Case "FindFilesResults" 'results
     msg = "Found " & .FindFilesResults.ListCount & ". Selected " & .FindFilesResults.SelCount & "."
     If morehelp = "" Then
     Else
        msg = msg & " Mouse on item " & morehelp
     End If
'     msg = msg & ".(Multiselection, DoubleClick to select all, RightClick for menu)"
  Case "Command1" 'start button
     msg = "Start search of " & .nametosearch.Text & " into selected directory"
  Case "Command2" 'abort/stop button
     msg = "Stop search of " & .nametosearch.Text & " into selected directory"
  Case "Drive1Label" 'drive1
     msg = "Select drive"
  Case "Dir1" 'dir1
     msg = "Select dir into which the search has to be done"
  Case "nametosearch" 'name to search
     msg = "Name to search (complete or partial, also wildcars allowed)"
  Case "frmAbout"
     msg = "The Author thanks you so much"
  Case "Pietro_Cecchi"
     msg = "Click here to send a message"
  Case "SaveCommand" 'on frmAbout
     msg = "Save this picture, at full size, as 'c:\Search About.bmp' "
  Case "PrintCommand" 'on frmAbout
     msg = "Print this picture, at the actual size unless resized, on default printer"
End Select

.HelpWindow.Panels(1).Text = msg

End With

End Sub

