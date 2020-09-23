Attribute VB_Name = "HSlistBox"
            
'***************************************************************
'Windows API/Global Declarations for :add a horizontal scroll bar
'     to a listbox or combo
'***************************************************************


#If Win16 Then

Public Declare Function SendMessage Lib "User" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
#Else

'Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


#End If


                        
 

'Source code:
'Note:This code is formatted to be pasted directly into VB.
'Pasting it into other editors may or may not work.


        
'***************************************************************
' Name: add a horizontal scroll bar to a listbox or combo
' Description:add a horizontal scroll bar to a listbox or combo b
'     ox
' By: VB Qaid
'
'
' Inputs:None
'
' Returns:None
'
'Assumes:None
'
'Side Effects:None
'
'Code provided by Planet Source Code(tm) (http://www.Planet-Sourc
'     e-Code.com) 'as is', without warranties as to performance, fitnes
'     s, merchantability,and any other warranty (whether expressed or i
'     mplied).
'This source code is copyrighted by Planet Source Code who has ex
'     clusive rights to distribute it.
'It is freely redistributable for personal use in source code for
'     m, or for personal or business use in a non-source code binary ex
'     ecutable.
'All other redistributions are prohibited without express written
'     consent from Exhedra Solutions, Inc.
'***************************************************************


'as in the original program
Public Const WM_USER = 1024
Public Const LB_SETHORIZONTALEXTENT = (WM_USER + 21)


'these are api constants
'Public Const WM_USER = &H400
'Public Const LB_SETHORIZONTALEXTENT = &H194


'this code in form:
'Dim nRet As Long
'Dim nNewWidth As Integer
'nNewWidth = List1.Width + 100 'new width in pixels
'nRet = SendMessage(List1.hwnd, LB_SETHORIZONTALEXTENT, nNewWidth, ByVal 0&)

 


