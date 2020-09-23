VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Intellisense Imitator -- Press the '<' key"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7830
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   353
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   522
   StartUpPosition =   2  'CenterScreen
   Begin IntellIm.ListSearch lsMain 
      Height          =   1815
      Left            =   1320
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   3201
   End
   Begin VB.TextBox txtMain 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   7815
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsLower 
         Caption         =   "&Lower Case"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuOptionsUpper 
         Caption         =   "&Upper Case"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuOptionsBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptionsComplete 
         Caption         =   "&Complete Tag"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************
' Project:  IntellIm                                         *
' Filename: frmMain.frm                                      *
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

Private Declare Function GetCaretPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Const vbKeyLessThan = 60

Dim pt          As POINTAPI
Dim lngStart    As Long
'

Private Sub Form_Load()
    
    mnuOptionsUpper.Checked = False
End Sub

Private Sub Form_Resize()
    
    txtMain.Move 0, 0, ScaleWidth, ScaleHeight
    
    ' This line of code moves the list view when the screen resolution changes
    If lsMain.Visible = True Then txtMain_KeyPress Val("<")
End Sub

Private Sub lsMain_Done(ByVal Text As String)
    
    ' Hide the popup window and add the text
    
    If mnuOptionsComplete.Checked = True Then
        
        ' Add the tag and close it
        txtMain.SelText = Text & "></" & Text & ">"
        
        ' Move the caret in between the two tags
        txtMain.SelStart = txtMain.SelStart - Len("</" & Text & ">")
    Else
        
        ' Add the tag without closing it
        txtMain.SelText = Text & ">"
    End If
        
    lsMain.Visible = False
    txtMain.SetFocus
End Sub

Private Sub lsMain_Escape()
    
    ' Hide the popup window and dont add the text
    lsMain.Visible = False
    txtMain.SetFocus
End Sub

Private Sub mnuFileExit_Click()
    
    Unload Me
End Sub

Private Sub mnuHelpAbout_Click()
    
    MsgBox "Intellisense Imitator" & vbCrLf & vbCrLf & "Copyright © 2000 Edward P. Denninger III" & vbCrLf & "edward3@optonline.net", vbInformation, "About"
End Sub

Private Sub mnuOptionsComplete_Click()
    
    mnuOptionsComplete.Checked = Not mnuOptionsComplete.Checked
End Sub

Private Sub mnuOptionsLower_Click()
    
    mnuOptionsLower.Checked = Not mnuOptionsLower.Checked
    mnuOptionsUpper.Checked = Not mnuOptionsUpper.Checked
End Sub

Private Sub mnuOptionsUpper_Click()
    
    mnuOptionsLower_Click
End Sub

Private Sub txtMain_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyLessThan Then
        
        ' Get the position of the caret
        GetCaretPos pt
        
        ' Get the selstart
        lngStart = txtMain.SelStart
        
        ' Move the popup window to the caret
        lsMain.Move pt.X + txtMain.Font.Size, pt.Y + (2 * txtMain.Font.Size)
        
        ' Check if the popup window is within the form
        If lsMain.Left + lsMain.Width > ScaleWidth Then lsMain.Move pt.X - lsMain.Width
        If lsMain.Top + lsMain.Height > ScaleHeight Then lsMain.Move lsMain.Left, pt.Y - lsMain.Height
        
        ' Fill the popup window with tags (only if there are no errors!)
        If lsMain.FillWithTags(App.Path & "\tags.lst", mnuOptionsUpper.Checked) = 0 Then Exit Sub
        
        ' Fill the popup window with fonts
        'lsMain.FillWithFonts
        
        ' Fill the popup window with available drives
        'lsMain.FillWithDrives mnuOptionsUpper.Checked
        
        ' Show the popup window
        lsMain.Visible = True
        
        ' Give the window focus
        lsMain.SetFocus
    End If
End Sub

Private Sub txtMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ' Hide the popup window
    lsMain.Visible = False
End Sub
