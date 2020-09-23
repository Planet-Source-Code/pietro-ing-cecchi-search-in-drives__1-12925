Attribute VB_Name = "Search"
Option Explicit

'Find File declarations and types
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * 260
        cAlternate As String * 14
End Type
'
'Convert Time Declare and type
'Public Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
'Public Type SYSTEMTIME
'        wYear As Integer
'        wMonth As Integer
'        wDayOfWeek As Integer
'        wDay As Integer
'        wHour As Integer
'        wMinute As Integer
'        wSecond As Integer
'        wMilliseconds As Integer
'End Type


Public SearchInterrupted As Boolean

Public Function FindAllFiles(Directory As String, Optional SearchFor As String)
    
    Dim Exists As Long
    Dim hFindFile As Long
    Dim FileData As WIN32_FIND_DATA
    
  With SearchForm
    
    'Sets Exists to equal 1
    'You need this so the loop doesn't automatically exit
    
    Exists = 1
    
    'Makes sure theres a "\" at the end of the directory
    If Right(Directory, 1) <> "\" Then Directory = Directory & "\"
    
    'Sets the default search item to *.*
    If SearchFor = vbNullString Then SearchFor = "*.*"
    
    'If the search for text doesn't contain any * or ?
    'Add *'s before and after
    If InStr(1, SearchFor, "?") = 0 And InStr(1, SearchFor, "*") = 0 Then
        SearchFor = "*" & SearchFor & "*"
    End If
    
    'Finds the first file
    hFindFile = FindFirstFile(Directory & SearchFor, FileData)
    
    Do While hFindFile <> -1 And Exists <> 0
        'A loop until all the files have been added
        
        DoEvents
        
        If (GetAttr(Directory & ClearNull(FileData.cFileName)) And vbDirectory) _
        = vbDirectory Then
            'If the file IS a directory than add it
            'to the temp listbox with the prefix DIR
            'don't list "." and ".." dirs
            If (ClearNull(FileData.cFileName) <> ".") And (ClearNull(FileData.cFileName) <> "..") Then
               .FindFilesTmpResults.AddItem "[dir] " & Directory & ClearNull(FileData.cFileName)
            End If
        ElseIf (GetAttr(Directory & ClearNull(FileData.cFileName)) And vbDirectory) _
        <> vbDirectory Then
            'If the file ISN'T a directory than add it
            'to the temp listbox with the prefix FILE
            .FindFilesTmpResults.AddItem "[file] " & Directory & ClearNull(FileData.cFileName)
        
        End If
         
        If SearchInterrupted Then 'interrupted by user
          Exit Function
        End If
        
        'Finds next file
        Exists = FindNextFile(hFindFile, FileData)
    Loop
    
    Do While .FindFilesTmpResults.ListCount
        'Removes everything from the temp listbox (Which is
        'alphabetically sorted, and puts it into the Viewed
        'Listbox
        'This is done so all the files are sorted alphabetically
        .FindFilesResults.AddItem .FindFilesTmpResults.List(0)
        .FindFilesTmpResults.RemoveItem 0
    Loop
    
    'Sets Exists to equal 1
    'You need this so the loop doesn't automatically exit
    Exists = 1
    
    'Find first file, this time includes directories in
    'the search
    hFindFile = FindFirstFile(Directory & "*", FileData)
    
    Do While hFindFile <> -1 And Exists <> 0
        'A loop until all the files have been added
      On Error GoTo skiptonextfile
        If (GetAttr(Directory & ClearNull(FileData.cFileName)) And vbDirectory) _
        = vbDirectory And (ClearNull(FileData.cFileName) <> "." And ClearNull(FileData.cFileName) <> "..") Then
           'If the file IS a directory and isn't "." or ".."
             'than adds it to the temp dir listbox
            
            .FindFilesTmpDirs.AddItem Directory & ClearNull(FileData.cFileName)
            DoEvents
        End If
nextfile:
      On Error GoTo 0
        
        
        'Finds next file
        Exists = FindNextFile(hFindFile, FileData)
    Loop

  End With
  
  Exit Function
  
skiptonextfile:
 Err.Clear
 Resume nextfile


End Function

Public Function ClearNull(StringToClear As String) As String
    Dim StartOfNulls As Long
    
    'This function clears all the nulls in the string and
    'Returns it, by using Instr to find the first null
    
    StartOfNulls = InStr(1, StringToClear, Chr(0))
    ClearNull = Left(StringToClear, StartOfNulls - 1)
End Function





Public Sub SearchFilesInDir(ByVal Directory As String, Optional SearchFor As String)

 Dim NextDir As String
    
  
  
  With SearchForm
    
    
    'Clears the result listbox
    .FindFilesResults.Clear
    .FindFilesTmpResults.Clear
    
    SearchInterrupted = False
    .SearchInterruptedLabel.Visible = False
    
    
    'Calls the FindAllFiles function
    FindAllFiles Directory, SearchFor

    Do While .FindFilesTmpDirs.ListCount
        'Searches through all the new directories and removes
        'Them from the temp dir listbox
        
        DoEvents
        
         
        NextDir = .FindFilesTmpDirs.List(0)
        .FindFilesTmpDirs.RemoveItem 0
        FindAllFiles NextDir, SearchFor
        
        If SearchInterrupted Then
           .FindFilesTmpDirs.Clear
           .FindFilesTmpResults.Clear
           .FindFilesResults.AddItem "  WARNING!: Search interrupted by user."
           .SearchInterruptedLabel.Caption = "Search Interrupted by user"
           .SearchInterruptedLabel.Visible = True
           Exit Sub
        End If
       
        'I put here a limit of 32000 items
        If .FindFilesResults.ListCount > 32000 Then '32767 is max short (ListCount)
            'Makes sure there aren't too many results
            'If there are too many, the listbox can't hold them
            
            .FindFilesTmpDirs.Clear
            .FindFilesTmpResults.Clear
            .FindFilesResults.AddItem "  WARNING!: Not all files have been listed. Only up to 7500 can be listed by this program."
            SearchInterrupted = True
            .SearchInterruptedLabel.Caption = "Search Interrupted by program"
            .SearchInterruptedLabel.Visible = True
            
            Exit Sub
        End If
        
    Loop
    

  End With
End Sub

