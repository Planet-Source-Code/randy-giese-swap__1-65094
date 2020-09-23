VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form RG_Help 
   BackColor       =   &H00FFFFE0&
   Caption         =   "Help  --->  www.RandyGrams.com"
   ClientHeight    =   3210
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   5250
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "RG_Help.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3210
   ScaleWidth      =   5250
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser HTML_Viewer 
      Height          =   2295
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   4575
      ExtentX         =   8070
      ExtentY         =   4048
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "RG_Help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'   *************************************************************************
'   *************************************************************************
'   ****                                                                 ****
'   ****    Display the HTML Help screen!                                ****
'   ****    www.RandyGrams.com                                           ****
'   ****                                                                 ****
'   *************************************************************************
'   *************************************************************************
'
Private RG_HTML_Help_File    As String

Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

Private Sub Form_Activate()

'   Place the Mouse Cursor on the "X" (Close)
    SetCursorPos (Screen.Width \ Screen.TwipsPerPixelX) - 13, 8

End Sub

Private Sub Form_Load()

'   Maximize the viewing area and open the Help file.
    With HTML_Viewer
        .Move 0, 0, Screen.Width, Screen.Height
        .Navigate Help_File
    End With      '  HTML_Viewer

End Sub

Private Sub Form_Unload(Cancel As Integer)

'   Unload this form before returning to the Main form.
    Unload Me

End Sub

Public Property Get Help_File() As String
'
'   *************************************************************************
'   *************************************************************************
'   ****                                                                 ****
'   ****    The Property Get/Let Procedures were suggested by Roger      ****
'   ****    Gilchrist to eliminate a Module that contained a single      ****
'   ****    Public Variable.  Thank you for the Help and other           ****
'   ****    suggestions Roger...  May 18, 2006                           ****
'   ****                                                                 ****
'   *************************************************************************
'   *************************************************************************
'
    Help_File = RG_HTML_Help_File

End Property

Public Property Let Help_File(ByVal FName As String)

    RG_HTML_Help_File = FName

End Property
