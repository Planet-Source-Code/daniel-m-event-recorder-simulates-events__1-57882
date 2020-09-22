VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmES 
   Caption         =   "Event Simulator"
   ClientHeight    =   1740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5220
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   1740
   ScaleWidth      =   5220
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkOverwrite 
      Caption         =   "Overwrite File"
      Height          =   255
      Left            =   3720
      TabIndex        =   6
      Top             =   720
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse..."
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   210
      Width           =   855
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   3720
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrPlay 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2640
      Top             =   600
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play Mouse Movement"
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox txtAcc 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Text            =   "1"
      Top             =   600
      Width           =   615
   End
   Begin VB.Timer tmrRecord 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3120
      Top             =   600
   End
   Begin VB.TextBox txtLoc 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   4095
   End
   Begin VB.CommandButton cmdRecord 
      Caption         =   "Record Mouse Movement"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lblRecordInt 
      Caption         =   "Record Interval (ms)"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   630
      Width           =   1575
   End
End
Attribute VB_Name = "frmES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'========================================================================
'Author: Daniel M.
'Submitted to: Planet Source Code (http://www.planet-source-code.com
'Comments: I didn't really comment the code, just provided the idea and
'the code works as described. Future plans is the ability to press
'multiple keys at once. Any comments or suggestions are appreciated.
'I also know that there may be bugs or things to work out but just
'wanted to share this with you. Note that some things may be commented out
'and are just old/unused code in the application and can be ignored.
'Enjoy the code and please vote for me! =)
'Contact: AIM: xAznHangukBoix             EMAIL: seoulxkorean@yahoo.com
'========================================================================
Dim PlayMovement As String
Dim bShift As Boolean
Dim bCaps As Boolean
Dim splitLines() As String
Dim strKeyPress As String
Dim skipProc As Boolean
Dim ac As Long, alreadyPressed(2) As Byte

Private Function funcKeyTrue(VirtualKey As Integer) As Boolean
'Handles mouse button presses
  On Error GoTo ERR:

         If GetKeyState(VirtualKey) = -127 Or _
            GetKeyState(VirtualKey) = -128 Then
        
                  funcKeyTrue = True
         End If
Exit Function
ERR:
   If ERR.Number <> 0 Then
       MsgBox ERR.Number & vbCrLf & ERR.Description
   End If
End Function

Private Sub cmdBrowse_Click()
With cdlg 'Open file for extended recording or over-write.
    .FileName = App.Path & "\Recorded\Recorded.esf"
    .Filter = "Event Simulator File (*.esf)|*.esf|All Files (*.*)|*.*|"
    .ShowOpen
End With

If Not vbCancel Then
    txtLoc.Text = cdlg.FileName
End If
End Sub

Private Sub cmdPlay_Click()
'=======================================================================
'COMMENTS: It can be noted that the text file could easily contain more
'information than the string can hold, therefore not allowing extended
'recording times. Future release will handle this by creating an array
'w/ the PlayMovement variable and will loop through that and the split
'lines to "play" the recording.
'=======================================================================
If cmdPlay.Caption = "Play Mouse Movement" Then
    Dim tempstr As String
    Open txtLoc.Text For Input As #1 'Open recorded information and while not end of file
        Do While Not EOF(1)
            Input #1, tempstr$ 'add the string to playmovement variable
            PlayMovement = PlayMovement & tempstr$ & vbNewLine
        Loop
    Close #1
    splitLines = Split(PlayMovement, vbNewLine, -1, 1) 'split up the string for each event
    ac = 0
    tmrPlay.Interval = txtAcc.Text
    tmrPlay.Enabled = True
    cmdPlay.Caption = "Stop Playing"
Else
    tmrPlay.Enabled = False
    cmdPlay.Caption = "Play Mouse Movement"
End If

End Sub

Private Sub cmdRecord_Click()
If cmdRecord.Caption = "Record Mouse Movement" Then
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FileExists(txtLoc.Text) Then
        fso.CreateTextFile (txtLoc.Text) 'Create text file if it cannot be located to hold information
        Else
        If chkOverwrite.Value = 1 Then 'If overwriting the file, delete the old one and create a new one.
            Kill txtLoc.Text
            fso.CreateTextFile (txtLoc.Text)
        Else
            If MsgBox("Do you wish to continue recording from previous file?", vbYesNo, "Continue Recording?") = vbYes Then
            
                Else
                Exit Sub
            End If
        End If
    End If
    alreadyPressed(0) = 0
    alreadyPressed(1) = 0
    
    tmrRecord.Interval = txtAcc.Text
    
    tmrRecord.Enabled = True

    
    cmdRecord.Caption = "Stop Record"
Else
    tmrRecord.Enabled = False

    
    PlayMovement = vbNullString
    cmdRecord.Caption = "Record Mouse Movement"
End If
End Sub

Private Sub Form_Load()
txtLoc.Text = App.Path & "\Recorded\Recorded.esf"
bShift = False
skipProc = False
End Sub

Private Sub tmrPlay_Timer()
'=====================================================================
'COMMENTS: Contains all the code for handling the event information
'from textfile. Currently handles left/right mouse button presses and
'most key presses, though not all at the moment.
'=====================================================================
If GetAsyncKeyState(vbKeyControl) And GetAsyncKeyState(vbKeyEnd) Then
    Call cmdPlay_Click
End If

Dim splitXY() As String
    If ac = UBound(splitLines) Then
        tmrPlay.Enabled = False
        cmdPlay.Caption = "Play Mouse Movement"
        Exit Sub
    End If
    splitXY = Split(splitLines(ac), "-", -1, 1)
    If UBound(splitXY) >= 3 Then
        Select Case splitXY(2) & splitXY(3)
        Case "RIGHTBUTTON"
            SetCursorPos splitXY(0), splitXY(1)
            mouse_event MOUSEEVENTF_RIGHTDOWN, splitXY(0), splitXY(1), 0, 0
        Case "RIGHTBUTTONUP"
            SetCursorPos splitXY(0), splitXY(1)
            mouse_event MOUSEEVENTF_RIGHTUP, splitXY(0), splitXY(1), 0, 0
            
        Case "LEFTBUTTON"
            SetCursorPos splitXY(0), splitXY(1)
            mouse_event MOUSEEVENTF_LEFTDOWN, splitXY(0), splitXY(1), 0, 0

        Case "LEFTBUTTONUP"
            SetCursorPos splitXY(0), splitXY(1)
            mouse_event MOUSEEVENTF_LEFTUP, splitXY(0), splitXY(1), 0, 0
            
        Case "KEYS"
            If splitXY(4) <> "COMMA" Then
                SendKeys splitXY(4), True
            Else
                SendKeys ",", True
            End If
        Case "KEYSPACE"
            SendKeys " ", True
        
        Case "KEYSUB"
            SendKeys "-", True
            
        Case "KEYB"
            SendKeys "{" & splitXY(4) & "}", True
        End Select
    Else
        SetCursorPos splitXY(0), splitXY(1)
    End If
    ac = ac + 1
End Sub

Private Sub tmrRecord_Timer()
'=============================================================
'COMMENTS: I know my method for handling mouse-down and mouse-
'up events is not very well-scripted but by the time I went
'back and looked at it I already had a lot done, so feel free
'to give me ideas on bettering this program.
'=============================================================
If GetAsyncKeyState(vbKeyControl) And GetAsyncKeyState(vbKeyEnd) Then 'hot key for ending recording
    Call cmdRecord_Click
End If

GetCursorPos Cursor

If alreadyPressed(0) = 1 And funcKeyTrue(VK_RBUTTON) = False Then
    strKeyPress = "RIGHT-BUTTONUP"
End If

If alreadyPressed(0) = 1 And funcKeyTrue(VK_RBUTTON) = True Then
    Else
    If funcKeyTrue(VK_RBUTTON) = True Then
        strKeyPress = "RIGHT-BUTTON"
        alreadyPressed(0) = 1
        Else
        alreadyPressed(0) = 0
    End If
End If

If alreadyPressed(1) = 1 And funcKeyTrue(VK_LBUTTON) = False Then
    strKeyPress = "LEFT-BUTTONUP"
End If

If alreadyPressed(1) = 1 And funcKeyTrue(VK_LBUTTON) = True Then
    Else
    If funcKeyTrue(VK_LBUTTON) = True Then
        strKeyPress = "LEFT-BUTTON"
        alreadyPressed(1) = 1
        Else
        alreadyPressed(1) = 0
    End If
End If

'=========================================================================
'NO SPECIAL REQUIREMENTS
If GetAsyncKeyState(&H6A) = -32767 Then strKeyPress = "KEY-S-*"
'If GetAsyncKeyState(&HBC) = -32767 Then strKeyPress = "KEY-S-,"
'If GetAsyncKeyState(&HBE) = -32767 Then strKeyPress = "KEY-S-."

'SPECIAL REQUIREMENTS
If GetAsyncKeyState(VK_Ctrl) = -32767 Then strKeyPress = "CTRL"
If GetAsyncKeyState(VK_Alt) = -32767 Then strKeyPress = "ALT"
If GetAsyncKeyState(VK_Del) = -32767 Then strKeyPress = "KEY-B-DEL"
If GetAsyncKeyState(&H20) = -32767 Then strKeyPress = "KEY-SPACE"

If GetAsyncKeyState(&H1B) = -32767 Then strKeyPress = "KEY-B-ESC"
If GetAsyncKeyState(&H21) = -32767 Then strKeyPress = "KEY-B-PGUP"
If GetAsyncKeyState(&H22) = -32767 Then strKeyPress = "KEY-B-PGDN"
If GetAsyncKeyState(&H23) = -32767 Then strKeyPress = "KEY-B-END"
If GetAsyncKeyState(&H24) = -32767 Then strKeyPress = "KEY-B-HOME"
If GetAsyncKeyState(&H25) = -32767 Then strKeyPress = "KEY-B-LEFT"
If GetAsyncKeyState(&H26) = -32767 Then strKeyPress = "KEY-B-UP"
If GetAsyncKeyState(&H27) = -32767 Then strKeyPress = "KEY-B-RIGHT"
If GetAsyncKeyState(&H28) = -32767 Then strKeyPress = "KEY-B-DOWN"
If GetAsyncKeyState(&HD) = -32767 Then strKeyPress = "KEY-B-ENTER"

If GetAsyncKeyState(&H2C) = -32767 Then strKeyPress = "KEY-B-PRNTSC"
If GetAsyncKeyState(&H8) = -32767 Then strKeyPress = "KEY-B-BKSP"
If GetAsyncKeyState(&H13) = -32767 Then strKeyPress = "KEY-B-BREAK"
If GetAsyncKeyState(&H23) = -32767 Then strKeyPress = "KEY-B-PGUP"
If GetAsyncKeyState(&H2D) = -32767 Then strKeyPress = "KEY-B-INS"

If GetAsyncKeyState(&H6B) = -32767 Then strKeyPress = "KEY-B-ADD"
If GetAsyncKeyState(&H6F) = -32767 Then strKeyPress = "KEY-B-DIVIDE"

If GetAsyncKeyState(&H90) = -32767 Then strKeyPress = "KEY-B-NUMLOCK"
If GetAsyncKeyState(&H91) = -32767 Then strKeyPress = "KEY-B-SCROLLLOCK"

'If GetAsyncKeyState(VK_LWIN) = -32767 Then strKeyPress = "KEY-B-LWIN"
'==========================================================================


For i = 112 To 127 'Function keys
    If GetAsyncKeyState(i) = -32761 Then
        strKeyPress = "KEY-B-F" & i - 111
    End If
Next i

If GetAsyncKeyState(20) = -32767 Or GetAsyncKeyState(20) = -32768 Then 'caps lock..
    bCaps = True
    Else
    bCaps = False
End If

If GetAsyncKeyState(16) = -32767 Or GetAsyncKeyState(16) = -32768 Then 'get shift state
    bShift = True
End If

If bShift = True Then
    For i = 65 To 90
        If GetAsyncKeyState(i) = -32767 Then 'alpha chars
            If bCaps = True Then strKeyPress = "KEY-S-" & Chr(i + 32) Else strKeyPress = "KEY-S-" & Chr(i)
        End If
    Next i
    
    For i = 48 To 57
        If GetAsyncKeyState(i) = -32767 Then 'number alts
                If i = 49 Then strKeyPress = "KEY-S-" & Chr(33) '!
                If i = 50 Then strKeyPress = "KEY-S-" & Chr(64) '@
                If i = 51 Then strKeyPress = "KEY-S-" & Chr(35) '#
                If i = 52 Then strKeyPress = "KEY-S-" & Chr(36) '$
                If i = 53 Then strKeyPress = "KEY-B-" & Chr(37) '%
                If i = 54 Then strKeyPress = "KEY-B-" & Chr(94) '^
                If i = 55 Then strKeyPress = "KEY-S-" & Chr(38) '&
                If i = 56 Then strKeyPress = "KEY-S-" & Chr(42) '*
                If i = 57 Then strKeyPress = "KEY-B-" & Chr(40) '(
                If i = 48 Then strKeyPress = "KEY-B-" & Chr(41) ')
        End If
    Next i
    
    'Misc chars
    If GetAsyncKeyState(&HC0) = -32767 Then strKeyPress = "KEY-S-~"
    If GetAsyncKeyState(&HBA) = -32767 Then strKeyPress = "KEY-S-:"
    If GetAsyncKeyState(&HBF) = -32767 Then strKeyPress = "KEY-S-?"
    If GetAsyncKeyState(&HBC) = -32767 Then strKeyPress = "KEY-S-<"
    If GetAsyncKeyState(&HBE) = -32767 Then strKeyPress = "KEY-S->"
    If GetAsyncKeyState(&HBD) = -32767 Then strKeyPress = "KEY-S-_"
    If GetAsyncKeyState(&HDC) = -32767 Then strKeyPress = "KEY-S-|"
    'If GetAsyncKeyState(&H2B) = -32767 Then strKeyPress = "KEY-B-+"
    If GetAsyncKeyState(&HDE) = -32767 Then strKeyPress = "KEY-S-" & Chr(34)
    
    If GetAsyncKeyState(&HDB) = -32767 Then strKeyPress = "KEY-B-{"
    If GetAsyncKeyState(&HDD) = -32767 Then strKeyPress = "KEY-B-}"
Else
     For i = 65 To 90
        If GetAsyncKeyState(i) = -32767 Then 'alpha chars
            If bCaps = True Then strKeyPress = "KEY-S-" & Chr(i) Else strKeyPress = "KEY-S-" & Chr(i + 32)
        End If
    Next i
    
    For i = 48 To 57 'numeric chars
        If GetAsyncKeyState(i) = -32767 Then
            strKeyPress = "KEY-S-" & Chr(i)
        End If
    Next i
    
    'Misc chars
    If GetAsyncKeyState(&HBC) = -32767 Then strKeyPress = "KEY-S-COMMA"
    If GetAsyncKeyState(&HBA) = -32767 Then strKeyPress = "KEY-S-;"
    If GetAsyncKeyState(&HBF) = -32767 Then strKeyPress = "KEY-S-/"
    If GetAsyncKeyState(&HC0) = -32767 Then strKeyPress = "KEY-S-`"
    'If GetAsyncKeyState(&H2B) = -32767 Then strKeyPress = "KEY-S-="
    If GetAsyncKeyState(&HDC) = -32767 Then strKeyPress = "KEY-S-\"
    If GetAsyncKeyState(&HDE) = -32767 Then strKeyPress = "KEY-S-'"
    If GetAsyncKeyState(&HBE) = -32767 Then strKeyPress = "KEY-S-."

    If GetAsyncKeyState(&HDB) = -32767 Then strKeyPress = "KEY-B-["
    If GetAsyncKeyState(&HDD) = -32767 Then strKeyPress = "KEY-B-]"
    If GetAsyncKeyState(&HBD) = -32767 Then strKeyPress = "KEY-SUB"

End If


'Print the positions along with any key combos
Open txtLoc.Text For Append As #1
    If strKeyPress <> vbNullString Then
        Print #1, Cursor.X & "-" & Cursor.Y & "-" & strKeyPress
    Else
        Print #1, Cursor.X & "-" & Cursor.Y
    End If
Close #1
bShift = False
strKeyPress = vbNullString
End Sub

Private Function Pause(dwMill As Long)
'===============================
'Pause Function
'===============================
Dim initTime As Long, fTime As Long
    initTime = GetTickCount 'get initial time
    Do Until fTime - initTime >= dwMill 'check time for pause function
        fTime = GetTickCount ' get current time
        DoEvents 'do events to prevent program from not working
    Loop
End Function

