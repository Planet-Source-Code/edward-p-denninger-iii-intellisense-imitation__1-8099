VERSION 5.00
Begin VB.UserControl ListSearch 
   ClientHeight    =   1695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1935
   ScaleHeight     =   113
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   129
   ToolboxBitmap   =   "ListSearch.ctx":0000
   Begin VB.TextBox txtSearch 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1935
   End
   Begin VB.ListBox lstMain 
      Height          =   1455
      IntegralHeight  =   0   'False
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "ListSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*************************************************************
' Project:  IntellIm                                         *
' Filename: ListSearch.ctl                                   *
' Author:   Edward P. Denninger III                          *
' Date:     5/14/2000                                        *
' Copyright © 2000 Edward P. Denninger III                   *
'*************************************************************
'*                         NOTICE                            *
'*************************************************************
' You may use and freely distribute this porject and source  *
' at your own leisure as long as I am given credit for my    *
' work.  If you have any comments or ideas for improvement,  *
' you can reach me at: edward3@optonline.net                 *
'*************************************************************

Option Explicit

' API Calls
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long


' API Constants
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Const LB_ADDSTRING = &H180
Private Const LB_FINDSTRING = &H18F
Private Const LB_RESETCONTENT = &H184
Private Const WS_DLGFRAME = &H400000


' Events
Public Event Done(ByVal Text As String)
Public Event Escape()
'

Public Function FillWithTags(ByVal Filename As String, Optional Uppercase As Boolean = True) As Integer
    On Error GoTo hErr
    
    Dim strTag As String
    
    ' Lock the window for a faster update
    LockWindowUpdate lstMain.hwnd
    
    ' Clear the listbox
    SendMessage lstMain.hwnd, LB_RESETCONTENT, 0&, ByVal 0&
    
    ' Fill the listbox with the tags from the file
    Open Filename For Input As #1
        Do Until EOF(1)
            Line Input #1, strTag
            
            If Uppercase Then
                
                ' If there is a tag then add it to the listbox
                If Len(strTag) > 0 Then SendMessage lstMain.hwnd, LB_ADDSTRING, 0&, ByVal UCase$(strTag)
            Else
                
                ' If there is a tag then add it to the listbox
                If Len(strTag) > 0 Then SendMessage lstMain.hwnd, LB_ADDSTRING, 0&, ByVal LCase$(strTag)
            End If
        Loop
    Close #1
    
    ' Unlock the window so we can see the tags
    LockWindowUpdate 0
    
    ' Return a value because we completed successfully
    FillWithTags = 1
hErr:
    Select Case Err.Number
    Case 0:
        Exit Function
    Case 53:
        MsgBox "Couldn't find the specified tags file!", vbExclamation, "Error"
        FillWithTags = 0
        Exit Function
    Case Else:
        MsgBox Err.Description, vbExclamation, "Error #" & Err.Number
        FillWithTags = 0
        Exit Function
    End Select
End Function

Public Sub FillWithFonts()
    Dim FontCounter As Long
    
    ' Lock the window for a faster update
    LockWindowUpdate lstMain.hwnd
    
    ' Clear the listbox
    SendMessage lstMain.hwnd, LB_RESETCONTENT, 0&, ByVal 0&
    
    ' Add the fonts
    For FontCounter = 0 To Screen.FontCount - 1
        
        SendMessage lstMain.hwnd, LB_ADDSTRING, 0&, ByVal Screen.Fonts(FontCounter) 'LCase$(strTag)
    Next FontCounter
    
    ' Unlock the window so we can see the fonts
    LockWindowUpdate 0
End Sub

Public Sub FillWithDrives(Optional Uppercase As Boolean = True)
    Dim strSave As String
    Dim ret     As Long
    Dim keer    As Integer
    
    ' Lock the window for a faster update
    LockWindowUpdate lstMain.hwnd
    
    ' Clear the listbox
    SendMessage lstMain.hwnd, LB_RESETCONTENT, 0&, ByVal 0&
    
    '--------------------------------------------------------------
    'KPD-Team 1998
    'URL: http://www.allapi.net/
    'E-Mail: KPDTeam@Allapi.net
    
    'Create a buffer to store all the drives
    strSave = String$(255, Chr$(0))
    
    'Get all the drives
    ret& = GetLogicalDriveStrings(255, strSave)
    
    'Extract the drives from the buffer and print them on the form
    For keer = 1 To 100
        If Left$(strSave, InStr(1, strSave, Chr$(0))) = Chr$(0) Then Exit For
        
        If Uppercase Then
            
            SendMessage lstMain.hwnd, LB_ADDSTRING, 0&, ByVal UCase$(Left$(strSave, InStr(1, strSave, Chr$(0)) - 1))
        Else
            
            SendMessage lstMain.hwnd, LB_ADDSTRING, 0&, ByVal LCase$(Left$(strSave, InStr(1, strSave, Chr$(0)) - 1))
        End If
        
        strSave = Right$(strSave, Len(strSave) - InStr(1, strSave, Chr$(0)))
    Next keer
    '--------------------------------------------------------------
    
    ' Unlock the window so we can see the drives
    LockWindowUpdate 0
End Sub

'---------------------------------------------------------
'- Control stuff -----------------------------------------
'---------------------------------------------------------
Private Sub lstMain_DblClick()
    
    lstMain_KeyPress (vbKeyReturn)
End Sub

Private Sub lstMain_KeyPress(KeyAscii As Integer)
    
    Select Case KeyAscii
    Case vbKeyReturn:
        
        RaiseEvent Done(lstMain.Text)
        txtSearch.Text = vbNullString
    Case vbKeyEscape:
        
        RaiseEvent Escape
        txtSearch.Text = vbNullString
    Case vbKeyDown:
        
        txtSearch.Text = lstMain.Text
    End Select
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    If KeyCode = vbKeyDown Then SetFocus lstMain.hwnd
End Sub

Private Sub txtSearch_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    Select Case KeyCode
    Case vbKeyReturn:
        
        RaiseEvent Done(txtSearch.Text)
        txtSearch.Text = vbNullString
    Case vbKeyEscape:
        
        RaiseEvent Escape
        txtSearch.Text = vbNullString
    Case Else:
        Dim lngListIndex As Long
        
        ' Get the list index of the item
        lngListIndex = SendMessage(lstMain.hwnd, LB_FINDSTRING, -1, ByVal txtSearch.Text)
        
        ' If the search string could not be found then...
        If lngListIndex = -1 Then
            
            Exit Sub
        Else    ' If the search string was found...
            
            lstMain.ListIndex = lngListIndex
        End If
    End Select
End Sub

Private Sub UserControl_Initialize()
    
    ' Put a raised border around the control
    SetWindowLong hwnd, GWL_STYLE, GetWindowLong(hwnd, GWL_STYLE) Or WS_DLGFRAME
End Sub

Private Sub UserControl_Paint()
    On Error Resume Next
    
    txtSearch.Move 0, 0, ScaleWidth
    lstMain.Move 0, txtSearch.Height + 1, ScaleWidth, ScaleHeight - lstMain.Top + 1
End Sub
