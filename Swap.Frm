VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form SwapMain 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Swap"
   ClientHeight    =   4935
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4380
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Swap.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   329
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   292
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picPause 
      Height          =   4815
      Left            =   60
      ScaleHeight     =   4755
      ScaleWidth      =   2940
      TabIndex        =   1
      Top             =   60
      Visible         =   0   'False
      Width           =   3000
      Begin VB.PictureBox picSolve 
         BackColor       =   &H00D8E9EC&
         ForeColor       =   &H00000000&
         Height          =   1335
         Left            =   1275
         ScaleHeight     =   1275
         ScaleWidth      =   1440
         TabIndex        =   14
         Top             =   3060
         Width           =   1500
         Begin VB.OptionButton optPauseTime 
            BackColor       =   &H00D8E9EC&
            Caption         =   "No"
            ForeColor       =   &H00000000&
            Height          =   420
            Index           =   12
            Left            =   307
            TabIndex        =   17
            Top             =   825
            Width           =   900
         End
         Begin VB.OptionButton optPauseTime 
            BackColor       =   &H00D8E9EC&
            Caption         =   "Yes"
            ForeColor       =   &H00000000&
            Height          =   420
            Index           =   11
            Left            =   307
            TabIndex        =   16
            Top             =   405
            Value           =   -1  'True
            Width           =   900
         End
         Begin VB.Label lblPause 
            AutoSize        =   -1  'True
            BackColor       =   &H00D8E9EC&
            Caption         =   "Solve now?"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   1
            Left            =   195
            TabIndex        =   15
            Top             =   120
            Width           =   1125
         End
      End
      Begin VB.OptionButton optPauseTime 
         BackColor       =   &H00D8E9EC&
         Caption         =   "1 sec"
         ForeColor       =   &H00000000&
         Height          =   420
         Index           =   10
         Left            =   120
         TabIndex        =   12
         Top             =   4320
         Width           =   900
      End
      Begin VB.OptionButton optPauseTime 
         BackColor       =   &H00D8E9EC&
         Caption         =   "9/10"
         ForeColor       =   &H00000000&
         Height          =   420
         Index           =   9
         Left            =   120
         TabIndex        =   11
         Top             =   3900
         Width           =   900
      End
      Begin VB.OptionButton optPauseTime 
         BackColor       =   &H00D8E9EC&
         Caption         =   "8/10"
         ForeColor       =   &H00000000&
         Height          =   420
         Index           =   8
         Left            =   120
         TabIndex        =   10
         Top             =   3480
         Width           =   900
      End
      Begin VB.OptionButton optPauseTime 
         BackColor       =   &H00D8E9EC&
         Caption         =   "7/10"
         ForeColor       =   &H00000000&
         Height          =   420
         Index           =   7
         Left            =   120
         TabIndex        =   9
         Top             =   3060
         Width           =   900
      End
      Begin VB.OptionButton optPauseTime 
         BackColor       =   &H00D8E9EC&
         Caption         =   "6/10"
         ForeColor       =   &H00000000&
         Height          =   420
         Index           =   6
         Left            =   120
         TabIndex        =   8
         Top             =   2640
         Width           =   900
      End
      Begin VB.OptionButton optPauseTime 
         BackColor       =   &H00D8E9EC&
         Caption         =   "5/10"
         ForeColor       =   &H00000000&
         Height          =   420
         Index           =   5
         Left            =   120
         TabIndex        =   7
         Top             =   2220
         Value           =   -1  'True
         Width           =   900
      End
      Begin VB.OptionButton optPauseTime 
         BackColor       =   &H00D8E9EC&
         Caption         =   "4/10"
         ForeColor       =   &H00000000&
         Height          =   420
         Index           =   4
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   900
      End
      Begin VB.OptionButton optPauseTime 
         BackColor       =   &H00D8E9EC&
         Caption         =   "3/10"
         ForeColor       =   &H00000000&
         Height          =   420
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   1380
         Width           =   900
      End
      Begin VB.OptionButton optPauseTime 
         BackColor       =   &H00D8E9EC&
         Caption         =   "2/10"
         ForeColor       =   &H00000000&
         Height          =   420
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   900
      End
      Begin VB.OptionButton optPauseTime 
         BackColor       =   &H00D8E9EC&
         Caption         =   "1/10"
         ForeColor       =   &H00000000&
         Height          =   420
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   540
         Width           =   900
      End
      Begin VB.OptionButton optPauseTime 
         BackColor       =   &H00D8E9EC&
         Caption         =   "0"
         ForeColor       =   &H00000000&
         Height          =   420
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   900
      End
      Begin VB.Label lblPause 
         Alignment       =   2  'Center
         BackColor       =   &H00D8E9EC&
         Caption         =   "This is the amount of time (in seconds) the program pauses while swapping boxes during the 'Solve' process."
         ForeColor       =   &H00FF0000&
         Height          =   2055
         Index           =   0
         Left            =   1080
         TabIndex        =   13
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.PictureBox picFrame 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1200
      Left            =   3120
      ScaleHeight     =   76
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   18
      Top             =   60
      Width           =   1200
      Begin VB.Label lblBox 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   450
         Index           =   0
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Visible         =   0   'False
         Width           =   450
      End
   End
   Begin MSComDlg.CommonDialog CDB 
      Left            =   3120
      Top             =   2580
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdChange 
      Appearance      =   0  'Flat
      BackColor       =   &H00C8D0D4&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4245
      TabIndex        =   0
      Top             =   2580
      Width           =   75
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Help!"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Options"
         Begin VB.Menu mnuType 
            Caption         =   "Type of Characters"
            Begin VB.Menu mnuChar 
               Caption         =   "Numbers"
               Index           =   1
               Begin VB.Menu mnuRand 
                  Caption         =   "Random"
                  Index           =   1
               End
               Begin VB.Menu mnuSeq 
                  Caption         =   "Sequential"
                  Index           =   1
               End
            End
            Begin VB.Menu mnuChar 
               Caption         =   "Upper Case Alpha"
               Index           =   2
               Begin VB.Menu mnuRand 
                  Caption         =   "Random"
                  Index           =   2
               End
               Begin VB.Menu mnuSeq 
                  Caption         =   "Sequential"
                  Index           =   2
               End
            End
            Begin VB.Menu mnuChar 
               Caption         =   "Lower Case Alpha"
               Index           =   3
               Begin VB.Menu mnuRand 
                  Caption         =   "Random"
                  Index           =   3
               End
               Begin VB.Menu mnuSeq 
                  Caption         =   "Sequential"
                  Index           =   3
               End
            End
            Begin VB.Menu mnuChar 
               Caption         =   "All Alpha"
               Index           =   4
            End
            Begin VB.Menu mnuChar 
               Caption         =   "WingDings Characters"
               Index           =   5
            End
            Begin VB.Menu mnuChar 
               Caption         =   "Special Characters"
               Index           =   6
            End
            Begin VB.Menu mnuChar 
               Caption         =   "All Ascii Characters"
               Index           =   7
            End
         End
         Begin VB.Menu mnuSize 
            Caption         =   "Puzzle Size"
            Begin VB.Menu mnuXbyX 
               Caption         =   "4*4"
               Index           =   1
            End
            Begin VB.Menu mnuXbyX 
               Caption         =   "5*5"
               Index           =   2
            End
            Begin VB.Menu mnuXbyX 
               Caption         =   "6*6"
               Index           =   3
            End
            Begin VB.Menu mnuXbyX 
               Caption         =   "7*7"
               Index           =   4
            End
            Begin VB.Menu mnuXbyX 
               Caption         =   "8*8"
               Index           =   5
            End
            Begin VB.Menu mnuXbyX 
               Caption         =   "9*9"
               Index           =   6
            End
         End
         Begin VB.Menu mnuRestart 
            Caption         =   "Restart"
         End
         Begin VB.Menu mnuSolve 
            Caption         =   "Solve!"
         End
         Begin VB.Menu mnuPause 
            Caption         =   "Solve Pause Time"
         End
      End
   End
End
Attribute VB_Name = "SwapMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'
'   *************************************************************************
'   *************************************************************************
'   ****                                                                 ****
'   ****    Swap:                                                        ****
'   ****    Written by:    Randy Giese   (03/25/02)                      ****
'   ****                                                                 ****
'   *************************************************************************
'   ****                                                                 ****
'   ****    Object:  Arrange the pieces so that no matching              ****
'   ****    Characters or Colors touch each other.                       ****
'   ****                                                                 ****
'   ****    It's simple to play but takes some logic to solve.           ****
'   ****                                                                 ****
'   *************************************************************************
'   *************************************************************************
'
'   RandyGrams Comments - Left Align the above comments.

Option Explicit

Private Type POINTAPI
    pntX                         As Long
    pntY                         As Long
End Type

Private Const SND_ASYNC          As Long = &H1        'play asynchronously
Private Const SND_FILENAME       As Long = &H20000    'name is a file name

Private blnSequential            As Boolean
Private blnSolve                 As Boolean
Private FSO                      As New Scripting.FileSystemObject

Private lngCol(1 To 100)         As Long
Private lngColor(1 To 100)       As Long
Private lngColors(33)            As Long
Private lngFntSz                 As Long
Private lngLetter                As Long
Private lngMsgResp               As Long
Private lngNumRows               As Long
Private lngPause                 As Long
Private lngPicTo                 As Long
Private lngRow(1 To 100)         As Long
Private lngSaveChar(1 To 9)      As Long
Private lngSaveColor(1 To 9)     As Long
Private lngSaveLoc               As Long
Private lngScrWidth              As Long
Private lngSlotNum(1 To 100)     As Long
Private lngSolveChar(1 To 9)     As Long
Private lngSolveColor(1 To 9)    As Long
Private lngSwap1                 As Long
Private lngSwap2                 As Long
Private lngTime                  As Long
Private lngType                  As Long
Private lngVal(1 To 100)         As Long

Private pntPoint                 As POINTAPI
Private strColor(33)             As String
Private strLetter                As String

'Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'Private Declare Function GetTickCount Lib "kernel32" () As Long
'Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
'Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

Private Sub CmdChange_Click()

    blnSolve = False
    SwapBoxes
    IsItAWinner

End Sub

Private Sub CreateBoxes()

Dim lngLoop1       As Long
Dim lngLoop2       As Long

    On Error Resume Next
    blnSolve = False
    strLetter = vbNullString

    With lblBox(0)
        .Visible = False
        .Font = "Times New Roman"
        .Height = picFrame.Height \ lngNumRows
        .Width = picFrame.Width \ lngNumRows
    End With    'lblBox(0)

    lngSaveLoc = 0
    LoadColors

    Select Case lngType
    Case 0 To 1              'Numeric
        lngLetter = 10
        For lngLoop1 = 1 To lngLetter - 1
            strLetter = strLetter & Trim$(Str$(lngLoop1))   'Fill the "Letter" array with Numbers
        Next lngLoop1
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    I purposely put the "0" last so that if the user chooses     ****
'   ****    "Sequential", the characters will be (1-2-3-4) instead of    ****
'   ****    (0-1-2-3).  If you prefer to start with "0", change the      ****
'   ****    "For" line (above) to:                                       ****
'   ****                                                                 ****
'   ****        For lngLoop1 = 0 To lngLetter                            ****
'   ****                                                                 ****
'   ****    then comment out, or delete:                                 ****
'   ****                                                                 ****
'   ****        strLetter = strLetter & "0"                              ****
'   ****                                                                 ****
'   *************************************************************************
'
'   RandyGrams Comments - Left Align the above comments.

        strLetter = strLetter & "0"
        Me.Caption = "Swap  'Numbers'"
    Case 2    'Upper Case Alpha
        lngLetter = 26
'       Fill the "Letter" array with Upper Case Letters
        For lngLoop1 = 1 To lngLetter
            strLetter = strLetter & Chr$(lngLoop1 + 64)
        Next lngLoop1

        Me.Caption = "Swap  'Upper Case Letters'"
    Case 3    'Lower Case Alpha
        lngLetter = 26
'       Fill the "Letter" array with Lower Case Letters
        For lngLoop1 = 1 To lngLetter
            strLetter = strLetter & Chr$(lngLoop1 + 96)
        Next lngLoop1

        Me.Caption = "Swap  'Lower Case Letters'"
    Case 4    'Upper and Lower Case Alpha
        lngLetter = 52
'       Fill the "Letter" array with Upper Case Letters
        For lngLoop1 = 1 To 26
            strLetter = strLetter & Chr$(lngLoop1 + 64)
        Next lngLoop1

'       Fill the "Letter" array with Lower Case Letters
        For lngLoop1 = 27 To lngLetter
            strLetter = strLetter & Chr$(lngLoop1 + 70)
        Next lngLoop1

        Me.Caption = "Swap  'All Letters'"
    Case 5    'Wingdings
        lngLetter = 223
'       Use the WingDings font
        lblBox(0).Font = "WingDings"
'       Fill the "Letter" array with Wingding Characters
        For lngLoop1 = 1 To lngLetter
            strLetter = strLetter & Chr$(lngLoop1 + 32)
        Next lngLoop1

'       Use the WingDings font
        lblBox(0).Font = "WingDings"
        Me.Caption = "Swap  'Wingding Characters'"
    Case 6 To 7 'Special Characters
'       Fill the "Letter" array with Special Characters
        For lngLoop1 = 33 To 255
            Select Case lngLoop1
            Case 33 To 47
                strLetter = strLetter & Chr$(lngLoop1)
            Case 58 To 64
                strLetter = strLetter & Chr$(lngLoop1)
            Case 91 To 96
                strLetter = strLetter & Chr$(lngLoop1)
            Case 123 To 126
                strLetter = strLetter & Chr$(lngLoop1)
            Case 128
                strLetter = strLetter & Chr$(lngLoop1)
            Case 130 To 140
                strLetter = strLetter & Chr$(lngLoop1)
            Case 142
                strLetter = strLetter & Chr$(lngLoop1)
            Case 145 To 156
                strLetter = strLetter & Chr$(lngLoop1)
            Case 158 To 159
                strLetter = strLetter & Chr$(lngLoop1)
            Case Is > 160
                strLetter = strLetter & Chr$(lngLoop1)
            End Select       'lngLoop1
        Next lngLoop1

        lngLetter = Len(strLetter)
        Me.Caption = "Swap  'Special Characters'"
    End Select               'lngType

    If lngType = 7 Then
        Me.Caption = "Swap  'All Ascii Characters'"
    End If

'   Create, Locate and Resize the boxes.
    For lngLoop1 = 1 To lngPicTo
        Load lblBox(lngLoop1)

        With lblBox(lngLoop1)
            .FontSize = lngFntSz
            .Visible = True
            .Move 0, 0, lblBox(0).Width \ Screen.TwipsPerPixelX, lblBox(0).Height \ Screen.TwipsPerPixelY
        End With         'lblBox(lngLoop1)
    Next lngLoop1

    PickColors

    For lngLoop1 = 1 To lngNumRows
        lngSaveColor(lngLoop1) = lngColor(lngLoop1 * lngNumRows)
        lngSaveChar(lngLoop1) = lngVal(lngLoop1)
        lngSwap1 = lngLoop1

'       Select & Draw the Characters
        DrawCharacters

        For lngLoop2 = 1 To lngNumRows
            lblBox(((lngLoop2 - 1) * lngNumRows) + lngLoop1).Caption = lblBox(lngLoop1).Caption
            lblBox(((lngLoop2 - 1) * lngNumRows) + lngLoop1).ToolTipText = lblBox(((lngLoop2 - 1) * lngNumRows) + lngLoop1).ToolTipText & "    Char# " & Asc(lblBox(lngLoop1).Caption)
        Next lngLoop2

        DoEvents             'Yield for other processing.
    Next lngLoop1

    lngTime = 3000
    Pause
    Rand

    On Error GoTo 0

End Sub

Private Sub DrawCharacters()

Dim lngTemp                     As Long
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    Draw the Characters within the boxes.                        ****
'   ****                                                                 ****
'   *************************************************************************
'
    If Not blnSequential Then
        lngTemp = Int(lngLetter * Rnd) + 1
    Else
        lngTemp = 1
    End If

    lblBox(lngSwap1).Caption = Mid$(strLetter, lngTemp, 1)
    lngLetter = lngLetter - 1
'   Remove the Letter that was just used from "strLetter"
    strLetter = Left$(strLetter, lngTemp - 1) & Trim$(Mid$(strLetter, lngTemp + 1, lngLetter))

End Sub

Private Sub Form_Load()

Dim FN1     As Long
Dim strType As String

    With SwapMain
        .Move 0, 0, Screen.Width, Screen.Height
        lngScrWidth = Screen.Width \ Screen.TwipsPerPixelX
        .ScaleMode = 3
        .Show
    End With  'SwapMain

    Select Case lngScrWidth
    Case 1280
        lngFntSz = 132
    Case 1024
        lngFntSz = 104
    Case 800
        lngFntSz = 82
    Case 640
        lngFntSz = 60
    End Select               'lngScrWidth

    picFrame.Move 5, 5, ScaleWidth - 10, ScaleHeight - 10
    picPause.Move 100, 100
'   The number of Rows & Columns
    lngNumRows = 4
'   How many Pics are needed
    lngPicTo = lngNumRows * lngNumRows
'   Set the Pause time during the Solve process.
    lngPause = 500

    If FSO.FileExists(App.Path & "\Swap.Opt") Then
        FN1 = FreeFile
        Open App.Path & "\Swap.Opt" For Input As #FN1
        Line Input #FN1, strType
        lngType = CLng(Trim$(Left$(strType, 1)))
        blnSequential = CBool(Right$(strType, 1))
        Close #FN1
    End If

    CreateBoxes

End Sub

Private Sub Form_Unload(Cancel As Integer)

Dim lngCntr                     As Long

    lngCntr = FreeFile
    Open App.Path & "\Swap.Opt" For Output As #lngCntr
    Print #lngCntr, Trim$(CStr(lngType)); Right$(CStr(CLng(blnSequential)), 1)
    Close #lngCntr
    Close

    For lngCntr = 2 To lngPicTo
        Unload lblBox(lngCntr)
    Next lngCntr

    Set SwapMain = Nothing
    End

End Sub

Private Sub HiLiteBox(ByVal lngIndex As Long)

'   *************************************************************************
'   ****                                                                 ****
'   ****    A box was clicked.  Make picture smaller so it stands out    ****
'   ****    and save it's location for the 'Swap' process.               ****
'   ****                                                                 ****
'   ****    If more than 2 boxes were clicked, reset the first one       ****
'   ****    back to original size.                                       ****
'   ****                                                                 ****
'   *************************************************************************
'
    lblBox(lngSwap2).Move lngCol(lngSwap2), lngRow(lngSwap2), lblBox(0).Width, lblBox(0).Height
    lngSwap2 = lngSwap1
    lngSwap1 = lngIndex
    lblBox(lngSwap1).Move lngCol(lngSwap1) + 5, lngRow(lngSwap1) + 5, lblBox(0).Width - 10, lblBox(0).Height - 10
    CmdChange.SetFocus

End Sub

Private Sub IsItAWinner()

    SeeIfSolved

    If Not blnSolve Then
        If lngSaveLoc = 0 Then
            PlaySound App.Path & "\Swap01.Wav", ByVal 0&, SND_FILENAME Or SND_ASYNC
            lngMsgResp = MsgBox("Congratulations!  Very good job!", vbOKOnly, "    Swap!    By:  Randy Giese")  'Display message.
            PlaySound App.Path & "\Swap02.Wav", ByVal 0&, SND_FILENAME
            PlaySound App.Path & "\Swap03.Wav", ByVal 0&, SND_FILENAME
            PlayAgain
        End If
    End If

End Sub

Private Sub lblBox_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim lngBakCol As Long
Dim lngInStr  As Long
Dim lngLoop1  As Long
Dim strLong   As String

    If Button = vbLeftButton Then
        HiLiteBox (Index)
    Else
        CDB.Flags = &H2
        lngBakCol = lblBox(Index).BackColor
        lblBox(Index).Move lngCol(Index) + 5, lngRow(Index) + 5, lblBox(0).Width - 10, lblBox(0).Height - 10
        CDB.ShowColor

        lblBox(Index).BackColor = CDB.Color
        lblBox(Index).Move lngCol(Index), lngRow(Index), lblBox(0).Width, lblBox(0).Height
        strLong = Right$("00000" & Hex$(CDB.Color), 6)
        strLong = "RGB (" & Mid$(strLong, 5, 2) & "," & Mid$(strLong, 3, 2) & "," & Left$(strLong, 2) & ")    "

'       Change the Color of all identically colored boxes.
        For lngLoop1 = 1 To lngPicTo
            With lblBox(lngLoop1)
                If .BackColor = lngBakCol Then
                    .BackColor = CDB.Color
                    lngInStr = InStr(1, .ToolTipText, "Char#", vbTextCompare)
                    .ToolTipText = strLong & Mid$(.ToolTipText, lngInStr)
                End If
            End With 'lblBox(lngLoop1)
        Next lngLoop1

        lblBox(Index).ToolTipText = strLong & Mid$(lblBox(Index).ToolTipText, lngInStr)
    End If

End Sub

Private Sub LoadColors()

'   *************************************************************************
'   *************************************************************************
'   ****                                                                 ****
'   ****    If you make any changes to these colors, remember to         ****
'   ****    change both the 'lngColors()' array and the 'strColor()'     ****
'   ****    array.                                                       ****
'   ****                                                                 ****
'   ****    Be sure to change the 'lngMax' variable in the               ****
'   ****    'PickColors' subRoutine if you add or delete any colors.     ****
'   ****    You may or may not have to change the dimensioning in the    ****
'   ****    'General Section'.  I set them both to '40' just in case     ****
'   ****    I wanted to add any more colors.                             ****
'   ****                                                                 ****
'   ****    The order of the colors is not important.  I just sorted     ****
'   ****    them by their long value.                                    ****
'   ****                                                                 ****
'   *************************************************************************
'   *************************************************************************
'
'   Load the Color array values
    lngColors(1) = 255         'Red
    lngColors(2) = 32768       'Dark Green
    lngColors(3) = 33023       'Orange
    lngColors(4) = 55295       'Gold
    lngColors(5) = 65280       'Lime
    lngColors(6) = 65535       'Yellow
    lngColors(7) = 2763429     'Brown
    lngColors(8) = 3381759     'Light Orange
    lngColors(9) = 6750207     'Light Yellow
    lngColors(10) = 8421504    'Gray
    lngColors(11) = 10040319    'Hot Pink
    lngColors(12) = 10079487    'Sand
    lngColors(13) = 10092543    'Chalk
    lngColors(14) = 13382297    'Dark Orchid
    lngColors(15) = 13408767    'Pink
    lngColors(16) = 13421823    'Faded Pink
    lngColors(17) = 13434828    'Ghost Green
    lngColors(18) = 13434879    'Pale Yellow
    lngColors(19) = 13882323    'Light Gray
    lngColors(20) = 14480885    'Beige
    lngColors(21) = 16118015    'Lavendar Blush
    lngColors(22) = 16443110    'Lavendar
    lngColors(23) = 16711680    'Blue
    lngColors(24) = 16711935    'Magenta
    lngColors(25) = 16724940    'Neon Purple
    lngColors(26) = 16737945    'Blue Violet
    lngColors(27) = 16751001    'Pastel Blue
    lngColors(28) = 16751052    'Easter Purple
    lngColors(29) = 16751103    'Light Violet
    lngColors(30) = 16764108    'Powder Blue
    lngColors(31) = 16775408    'Alice Blue
    lngColors(32) = 16776960    'Cyan
    lngColors(33) = 16777113    'Ice Blue

'   Load the Color array Names (used in the ToolTipText)
    strColor(1) = "Red"
    strColor(2) = "Dark Green"
    strColor(3) = "Orange"
    strColor(4) = "Gold"
    strColor(5) = "Lime"
    strColor(6) = "Yellow"
    strColor(7) = "Brown"
    strColor(8) = "Light Orange"
    strColor(9) = "Light Yellow"
    strColor(10) = "Gray"
    strColor(11) = "Hot Pink"
    strColor(12) = "Sand"
    strColor(13) = "Chalk"
    strColor(14) = "Dark Orchid"
    strColor(15) = "Pink"
    strColor(16) = "Faded Pink"
    strColor(17) = "Ghost Green"
    strColor(18) = "Pale Yellow"
    strColor(19) = "Light Gray"
    strColor(20) = "Beige"
    strColor(21) = "Lavendar Blush"
    strColor(22) = "Lavendar"
    strColor(23) = "Blue"
    strColor(24) = "Magenta"
    strColor(25) = "Neon Purple"
    strColor(26) = "Blue Violet"
    strColor(27) = "Pastel Blue"
    strColor(28) = "Easter Purple"
    strColor(29) = "Light Violet"
    strColor(30) = "Powder Blue"
    strColor(31) = "Alice Blue"
    strColor(32) = "Cyan"
    strColor(33) = "Ice Blue"

End Sub

Private Sub mnuAbout_Click()

'   Show the "About" screen.
    RG_About.Show vbModal, Me
'   When we return,
    SetCursorPos pntPoint.pntX, pntPoint.pntY

End Sub

Private Sub mnuChar_Click(Index As Integer)

    lngType = Index

    Select Case lngType
    Case 4 To 7
        blnSequential = False
        Restart
    End Select               'lngType

End Sub

Private Sub mnuHelp_Click()

'   The name of the Help file.
    RG_Help.Help_File = App.Path & "\Swap.Html"
'   Display the Help page
    RG_Help.Show vbModal, Me
'   When we return,
    SetCursorPos pntPoint.pntX, pntPoint.pntY

End Sub

Private Sub mnuMenu_Click()

'   Get and save the current cursor position
    GetCursorPos pntPoint

End Sub

Private Sub mnuPause_Click()

    picPause.Visible = True

End Sub

Private Sub mnuRand_Click(Index As Integer)

    blnSequential = False
    Restart

End Sub

Private Sub mnuRestart_Click()

    Restart

End Sub

Private Sub mnuSeq_Click(Index As Integer)

    blnSequential = True
    Restart

End Sub

Private Sub mnuSolve_Click()

    blnSolve = True
    SolveIt

End Sub

Private Sub mnuXbyX_Click(Index As Integer)

    lngNumRows = Index + 2
    lngMsgResp = 6
'
'   *************************************************************************
'   *************************************************************************
'   ****                                                                 ****
'   ****    The variable "lngFntSz" is the size of the font used         ****
'   ****    depending on the puzzle size.                                ****
'   ****                                                                 ****
'   ****    Note:  These font sizes were determined using a screen       ****
'   ****    resolution of 1024*768.  Other screen resolutions may not    ****
'   ****    display correctly.                                           ****
'   ****                                                                 ****
'   *************************************************************************
'   ****                                                                 ****
'   ****    Fixed on September 19, 2006.                                 ****
'   ****    The characters should display correctly for all puzzles      ****
'   ****    when the Screen resolution is one of the following:          ****
'   ****                                                                 ****
'   ****    1280*960                                                     ****
'   ****    1024*768                                                     ****
'   ****    800*600                                                      ****
'   ****    640*480                                                      ****
'   ****                                                                 ****
'   *************************************************************************
'   *************************************************************************
'
    Select Case lngScrWidth
    Case 1280
'       Screen resolution:  1280*960
        Select Case Index
        Case 1
            lngFntSz = 132
        Case 2
            lngFntSz = 112
        Case 3
            lngFntSz = 92
        Case 4
            lngFntSz = 80
        Case 5
            lngFntSz = 68
        Case 6
            lngFntSz = 60
        End Select           'Index
    Case 1024
'       Screen resolution:  1024*768
        Select Case Index
        Case 1
            lngFntSz = 104
        Case 2
            lngFntSz = 84
        Case 3
            lngFntSz = 68
        Case 4
            lngFntSz = 58
        Case 5
            lngFntSz = 50
        Case 6
            lngFntSz = 44
        End Select           'Index
    Case 800
'       Screen resolution:  800*600
        Select Case Index
        Case 1
            lngFntSz = 82
        Case 2
            lngFntSz = 66
        Case 3
            lngFntSz = 56
        Case 4
            lngFntSz = 44
        Case 5
            lngFntSz = 40
        Case 6
            lngFntSz = 36
        End Select           'Index
    Case 640
'       Screen resolution:  640*480
        Select Case Index
        Case 1
            lngFntSz = 60
        Case 2
            lngFntSz = 48
        Case 3
            lngFntSz = 40
        Case 4
            lngFntSz = 32
        Case 5
            lngFntSz = 30
        Case 6
            lngFntSz = 28
        End Select           'Index
    End Select               'lngScrWidth

    Restart

End Sub

Private Sub optPauseTime_Click(Index As Integer)

    picPause.Visible = False

    Select Case Index
        Case 0 To 10
            lngPause = Index * 100
    End Select  'Index

    If optPauseTime(11).Value = True Then
        mnuSolve_Click
    End If

End Sub

Private Sub Pause()

Dim lngStop                  As Long

    lngStop = GetTickCount + lngTime
    Do While GetTickCount < lngStop
        DoEvents
    Loop

End Sub

Private Sub PickColors()

Dim lngLoop1 As Long
Dim lngLoop2 As Long
Dim lngMax   As Long
Dim lngRCNum As Long
Dim lngRGB   As Long

    lngMax = UBound(lngColors())   'Highest element in the "lngColors" array

    For lngLoop1 = 1 To lngNumRows
        Randomize Timer
'       Pick a "lngColors" array element
'       lngRGB=lngLoop1  'This will always use the same colors.
        lngRGB = Int(lngMax * Rnd) + 1

        For lngLoop2 = 1 To lngNumRows
            lngRCNum = ((lngLoop1 - 1) * lngNumRows) + lngLoop2
            lngCol(lngRCNum) = lblBox(0).Width * (lngLoop2 - 1)
            lngRow(lngRCNum) = lblBox(0).Height * (lngLoop1 - 1)
            lngSlotNum(lngRCNum) = lngRCNum

            With lblBox(lngRCNum)
                .Visible = True
                .BackColor = lngColors(lngRGB)
                .ToolTipText = strColor(lngRGB)
                .Move lngCol(lngRCNum), lngRow(lngRCNum), lblBox(0).Width, lblBox(0).Height
            End With         'lblBox(lngRCNum)

            lblBox(lngRCNum).ToolTipText = strColor(lngRGB)
            lngColor(lngRCNum) = lngLoop1   'lngColor will be the same for the whole Row.
            lngVal(lngRCNum) = lngLoop2     'lngVal will be the same value for each Character in a column.
        Next lngLoop2
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    This will remove the color that was just used from the       ****
'   ****    "lngColors" array.                                           ****
'   ****                                                                 ****
'   *************************************************************************
'
'   RandyGrams Comments - Left Align the above comments.

        lngColors(lngRGB) = lngColors(lngMax)
        strColor(lngRGB) = strColor(lngMax)
        lngMax = lngMax - 1
    Next lngLoop1

End Sub

Private Sub PlayAgain()

    lngMsgResp = MsgBox("Would you like to try a more difficult screen?", vbYesNoCancel, "    Swap!    By:  Randy Giese")   'Display message.

    Select Case lngMsgResp
    Case vbYes
        lngFntSz = lngFntSz - 20
        Restart
    Case vbNo
        Restart
    End Select               'lngMsgResp

End Sub

Private Sub Rand()

Dim lngLoop3                    As Long
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    Scramble the boxes.                                          ****
'   ****                                                                 ****
'   *************************************************************************
'
    Randomize Timer

    For lngLoop3 = 1 To lngPicTo * lngNumRows
        lngSwap1 = Int(lngPicTo * Rnd) + 1
        lngSwap2 = Int(lngPicTo * Rnd) + 1
        SwapBoxes
    Next lngLoop3

End Sub

Private Sub Restart()

Dim lngLoop1                    As Long

    On Error Resume Next

'   Delete the "Old" Labels and Boxes
    For lngLoop1 = 1 To lngPicTo
        Unload lblBox(lngLoop1)
    Next lngLoop1

    If lngMsgResp = 6 Then
        lngMsgResp = 0
'       The number of Rows & Columns
        lngNumRows = lngNumRows + 1

        If lngNumRows > 9 Then
            lngNumRows = 9
            lngMsgResp = MsgBox("I'm sorry, but you have reached the current extent of this program." & vbCr & "Please do a Print Screen and E-mail a copy to me at:" & vbCr & "WebMaster@RandyGrams.com" & vbCr & "Thank you!", vbOKOnly, "    Swap!    By:  Randy Giese")
        End If

'       How many Pics are needed (starts at 1)
        lngPicTo = lngNumRows * lngNumRows
    End If

'   Reset the Cursor
    SetCursorPos pntPoint.pntX, pntPoint.pntY
    CreateBoxes

    On Error GoTo 0

End Sub

Private Sub SavedColors()

Dim lngLoop1       As Long

    For lngLoop1 = 1 To lngNumRows
        lngSolveColor(lngLoop1) = lngSaveColor(lngLoop1)
        lngSolveChar(lngLoop1) = lngSaveChar(lngLoop1)
    Next lngLoop1

End Sub

Private Sub SeeIfSolved()

Dim lngFlag  As Long
Dim lngLoop1 As Long
Dim lngLoop2 As Long

    For lngLoop1 = 1 To lngPicTo - 1
        For lngLoop2 = lngLoop1 + 1 To lngPicTo
            lngFlag = 0
            lngSaveLoc = 0

'           See if boxes are in the same or adjacent rows
            If Abs(((lngLoop2 - 1) \ lngNumRows) - ((lngLoop1 - 1) \ lngNumRows)) < 2 Then
                lngFlag = 1
            End If

'           See if boxes are in the same or adjacent columns
            If Abs((lngLoop2 - 1) Mod lngNumRows - (lngLoop1 - 1) Mod lngNumRows) < 2 Then
                lngFlag = lngFlag + 1
            End If

            Select Case lngFlag
            Case 0
'               If no boxes are touching, exit the loop
                Exit For
            Case 2
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    If 2 boxes are touching, check the Characters and the        ****
'   ****    Colors. If the Values are equal, the puzzle is not           ****
'   ****    solved. The Values will be equal if the Characters are       ****
'   ****    identical.                                                   ****
'   ****                                                                 ****
'   *************************************************************************
'
'   RandyGrams Comments - Left Align the above comments.

                If lngVal(lngSlotNum(lngLoop1)) = lngVal(lngSlotNum(lngLoop2)) Then
                    lngSaveLoc = lngLoop2   'lngSaveLoc saves the location of the first 'Out of Order' Character
                    Exit For
                End If

'               If the Colors are equal, the puzzle is not solved.
                If lngColor(lngSlotNum(lngLoop1)) = lngColor(lngSlotNum(lngLoop2)) Then
                    lngSaveLoc = lngLoop2   'lngSaveLoc saves the location of the first 'Out of Order' Character
                    Exit For
                End If
            End Select       'lngFlag

            DoEvents         'Yield for other processing.
        Next lngLoop2

        If lngSaveLoc Then
            Exit For
        End If
    Next lngLoop1

End Sub

Private Sub SolveIt()

Dim lngLoop1      As Long
Dim lngStart      As Long
Dim strSolveColor As String
Dim strSolveChar  As String

    For lngLoop1 = 1 To lngNumRows
        strSolveChar = strSolveChar & lngLoop1
    Next lngLoop1

    strSolveChar = strSolveChar & Trim$(Mid$(strSolveChar, (lngNumRows \ 2) + 1, lngNumRows + 1))
    strSolveChar = strSolveChar & Left$(strSolveChar, lngNumRows \ 2)

    For lngLoop1 = 1 To 3
        strSolveChar = strSolveChar & strSolveChar
    Next lngLoop1

    Select Case lngNumRows
    Case 4
        strSolveColor = "1212343421214343"
    Case 5
'       strSolveColor = "1212134343515154343425252"
        strSolveColor = "1513124242313534242453515"
        strSolveChar = "1234554123123455412312345"
    Case 6
        strSolveColor = "121212343434212121565656343434656565"
    Case 7
'       strSolveColor = "1212121343434351515156363636757575746464642727272"
        strSolveColor = "1212121343737456565657374743656565647434372121212"
        strSolveChar = "1234567765123412345677651234123456776512341234567"
    Case 8
        strSolveColor = "1212121234343434212121214343434356565656787878786565656587878787"
    Case 9
'        strSolveColor = "121212121343434343565656565787878787959595959878787878696969696434343434212121212"
        strSolveColor = "121212121343434343565956595787878787959695969878787878696569656434343434212121212"
    End Select  '   lngNumRows

    SavedColors
    lngStart = 1

    Do While lngPicTo > lngStart
        For lngLoop1 = lngStart To lngPicTo
            If lngColor(lngSlotNum(lngLoop1)) = lngSolveColor(Mid$(strSolveColor, lngStart, 1)) Then
                If lngVal(lngSlotNum(lngLoop1)) = lngSolveChar(Mid$(strSolveChar, lngStart, 1)) Then
                    lngSwap1 = lngStart
                    HiLiteBox (lngSwap1)
                    lngSwap2 = lngLoop1
                    HiLiteBox (lngSwap2)

                    lngTime = lngPause
                    Pause
                    SwapBoxes

                    lngStart = lngStart + 1
                    Pause
                    Exit For
                End If
            End If

            DoEvents
        Next lngLoop1
    Loop

    PlaySound App.Path & "\Tada.Wav", ByVal 0&, SND_FILENAME Or SND_ASYNC

End Sub

Private Sub SwapBoxes()

Dim lngSwap                     As Long
Dim strSwap                     As String

'   Swap the Colors
    strSwap = lblBox(lngSwap1).BackColor
    lblBox(lngSwap1).BackColor = lblBox(lngSwap2).BackColor
    lblBox(lngSwap2).BackColor = strSwap

'   Swap the Labels
    strSwap = lblBox(lngSwap1).Caption
    lblBox(lngSwap1).Caption = lblBox(lngSwap2).Caption
    lblBox(lngSwap2).Caption = strSwap

'   Swap the ToolTipText messages
    strSwap = lblBox(lngSwap1).ToolTipText
    lblBox(lngSwap1).ToolTipText = lblBox(lngSwap2).ToolTipText
    lblBox(lngSwap2).ToolTipText = strSwap

'   Reset the Pics back to original size
    lblBox(lngSwap1).Move lngCol(lngSwap1), lngRow(lngSwap1), lblBox(0).Width, lblBox(0).Height
    lblBox(lngSwap2).Move lngCol(lngSwap2), lngRow(lngSwap2), lblBox(0).Width, lblBox(0).Height

'   Update the Slot-Number Array (lngSlotNum)
    lngSwap = lngSlotNum(lngSwap1)
    lngSlotNum(lngSwap1) = lngSlotNum(lngSwap2)
    lngSlotNum(lngSwap2) = lngSwap

End Sub
