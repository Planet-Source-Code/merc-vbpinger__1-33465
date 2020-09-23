VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMain 
   Caption         =   "Pinger"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7545
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   7545
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   285
      Index           =   3
      Left            =   1680
      Picture         =   "frmMain.frx":49E2
      ScaleHeight     =   225
      ScaleWidth      =   450
      TabIndex        =   10
      Top             =   6525
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   285
      Index           =   2
      Left            =   1120
      Picture         =   "frmMain.frx":4F88
      ScaleHeight     =   225
      ScaleWidth      =   450
      TabIndex        =   9
      Top             =   6525
      Visible         =   0   'False
      Width           =   510
   End
   Begin MSFlexGridLib.MSFlexGrid grdDetails 
      Height          =   3855
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   6800
      _Version        =   393216
      Rows            =   3
      Cols            =   6
      FixedCols       =   0
      ScrollBars      =   2
   End
   Begin MSComctlLib.ImageList imgToolPics 
      Left            =   5760
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":552E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9F20
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F5B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FA04
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   285
      Index           =   1
      Left            =   560
      Picture         =   "frmMain.frx":14E9E
      ScaleHeight     =   225
      ScaleWidth      =   450
      TabIndex        =   6
      Top             =   6525
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   285
      Index           =   0
      Left            =   0
      Picture         =   "frmMain.frx":15444
      ScaleHeight     =   225
      ScaleWidth      =   450
      TabIndex        =   5
      Top             =   6525
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.ListBox lstErrors 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   120
      TabIndex        =   4
      Top             =   5325
      Width           =   7335
   End
   Begin MSComctlLib.StatusBar staStatus 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   6495
      Width           =   7545
      _ExtentX        =   13309
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrPing 
      Left            =   6360
      Top             =   0
   End
   Begin MSComctlLib.Toolbar tbrTools 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7545
      _ExtentX        =   13309
      _ExtentY        =   1376
      ButtonWidth     =   1191
      ButtonHeight    =   1376
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "imgToolPics"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Ping"
            Object.ToolTipText     =   "Start Pinging"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Hosts"
            Object.ToolTipText     =   "Add New Host"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Config"
            Object.ToolTipText     =   "Configure Program"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "E&xit"
            Object.ToolTipText     =   "Exit !"
            ImageIndex      =   4
         EndProperty
      EndProperty
      Begin MSComCtl2.Animation aniPing 
         Height          =   480
         Left            =   6960
         TabIndex        =   2
         Top             =   180
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         AutoPlay        =   -1  'True
         FullWidth       =   32
         FullHeight      =   32
      End
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Date/Time            Machine           Error"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   120
      TabIndex        =   7
      Top             =   5025
      Width           =   7335
   End
   Begin VB.Label Label1 
      Caption         =   "Errors:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   3
      Top             =   4785
      Width           =   1245
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add Machine"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************************
'*                                ICMP Ping Program                                    *
'*                               Last Updated 26/03/2002                               *
'*                                                                                     *
'* My first contribution to PSC. This program demonstrates how to use API based        *
'* calls to Ping one (or more) machines over a standard TCPIP connection.              *
'* It also demonstrates useage of MS Flex Grid, Animation Control and Toolbar controls *
'* supplied with MS Visual Basic aswell as a custom registry control module.           *
'* I've commented the code as best I can and hope you'll be able to understand it      *
'*                                                                                     *
'* I'd appreciate any feedback you have to offer, good or otherwise                    *                                                               *
'*                                                                                     *
'*       If you have any problems or queries: martin.sidgreaves@wrigley.co.uk          *
'***************************************************************************************

Option Explicit
Private ping_array(999) As Integer
Private list_count As Integer
Private timer_count As Long
Private response As String

Private Sub do_ping()

  Dim intLcount As Integer
  Dim strPing As String
  Dim intCount As Integer
  Dim strTmp As String
  Dim strWrite As String
  Dim lcount As Integer
  Dim objMail As Object
  Dim SQL As String
  Dim dteAdded As Date
  Dim intLogFile As Integer

    With grdDetails
        For intLcount = 2 To .Rows
            .Row = intLcount - 1
            .Col = 2
            strPing = .Text
            .Col = 3

            'Put a blue line in while pinging
            .Row = intLcount - 1
            .Col = 3
            .Text = "Testing"
            .CellBackColor = vbBlue
            .CellForeColor = vbWhite
            .Col = 4
            .Text = Now

            'Perform the ping
            strTmp = CStr(Ping(strPing))

            .Col = 5
            If .CellPicture = Picture1(2).Picture Then
                'Check to see if directory exists
                If Not DirExists("C:\Ping Results") Then
                    'It doesn't, create it!
                    MkDir ("C:\Ping Results")
                End If

                'Log this machine's data
                .Col = 1
                intLogFile = FreeFile
                Open "C:\Ping Results\" & .Text & ".txt" For Append As intLogFile
                Print #intLogFile, Now, strTmp
                Close #intLogFile
            End If

            If strTmp < 0 Then
                'If it's a bad one.... make it red
                .Col = 0
                .ColAlignment(0) = 1
                Set .CellPicture = Picture1(0).Picture

                .Col = 3
                .Row = intLcount - 1
                .Text = "Error"
                .CellBackColor = vbRed
                .CellForeColor = vbWhite

                'Log the error
                strWrite = CStr(Now) & "   "
                .Col = 1
                strWrite = strWrite & Rpad(.Text, " ", 18)
                strWrite = strWrite & "Timeout"

                'Add it to the array
                ping_array(intLcount) = ping_array(intLcount) + 1

                'Check for ping timeout
                If ping_array(intLcount) >= pings Then

                    'Start the animation
                    lResID = 102
                    aniPing.AutoPlay = True
                    LoadResAVI aniPing, lResID
                    
                    .Col = 4
                    .CellBackColor = vbRed
                    .CellForeColor = vbWhite
                    staStatus.Panels(4).Text = "Error: TimeOut"

                    'Reset array
                    ping_array(intLcount) = 0
                    .Col = 4
                    .CellBackColor = vbWhite
                    .CellForeColor = vbBlack
                    staStatus.Panels(4).Text = ""

                    'Write it to the log window
                    lstErrors.AddItem strWrite

                    'Log it to file too!
                    Print #fnum, strWrite
                End If
              ElseIf strTmp < 11 Then

                    'If it's a good response... make it green
                    .Col = 3
                   .Text = "< 10ms"
                   .Row = intLcount - 1
                   .Col = 0
                   Set .CellPicture = Picture1(1).Picture
                  .Col = 3
                  .CellForeColor = vbBlack
                  .CellBackColor = vbGreen
                  ping_array(intLcount) = 0

              ElseIf strTmp >= 11 Then

                  'If it's a good response... make it green
                  .Col = 3
                  .Text = strTmp & "ms"
                  .Row = intLcount - 1
                  .Col = 0
                  Set .CellPicture = Picture1(1).Picture
                  .Col = 3
                  .CellForeColor = vbBlack
                  .CellBackColor = vbGreen
                  ping_array(intLcount) = 0

            End If
        Next intLcount
    End With 'grdDetails

    'Play the animation
    lResID = 101
    aniPing.AutoPlay = True
    LoadResAVI aniPing, lResID

End Sub

Private Sub Form_Activate()

  Dim introw As Integer
  Dim intMfile As Integer
  Dim strIP As String
  Dim strName As String
  Dim strAlert As String
  Dim intFload As Integer
  Dim strVal As String
  Dim strLog As String

    On Error GoTo bad_file

    'Complete the status bar
    staStatus.Panels(2).Text = "Delay: " & CStr(delay) & " Mins"
    staStatus.Panels(3).Text = "Pings: " & CStr(pings)
     
    'Clear the listbox
    lstErrors.Clear
     
    'Set up the grid control
    introw = 1
    intMfile = FreeFile
    With grdDetails
        .Clear
        .Col = 1
        .Row = 0
        .Text = "Machine Name"
        .Col = 2
        .Row = 0
        .Text = "IP Address"
        .Col = 3
        .Row = 0
        .Text = "Response"
        .Col = 4
        .Row = 0
        .Text = "Last Ping"
        .Col = 5
        .Row = 0
        .Text = "Log"

        'Open the saved computer file
        Open App.Path & "\machines.dat" For Input As intMfile
        Do
            Input #intMfile, strName
            If Mid(strName, 1, 1) <> "#" Then
                Input #intMfile, strIP
                Input #intMfile, strLog
                .Col = 1
                .Rows = introw + 1
                .Row = introw
                .Text = strName
                .Col = 2
                .Row = introw
                .Text = strIP
                .Col = 5
                .Row = introw

                'Show whether logging is on or off for this machine
                If strLog = "1" Then
                    Set .CellPicture = Picture1(2).Picture
                  Else
                    Set .CellPicture = Picture1(3).Picture
                End If

                ping_array(introw) = 0
                list_count = .Rows
                introw = introw + 1
            End If
        Loop While Not EOF(intMfile)
        Close #intMfile
    End With 'grdDetails

    'Do an initial ping once the form has initialised
    DoEvents
    do_ping
    timer_count = 0
    tmrPing.Enabled = True

    'Play the animation
    lResID = 101
    aniPing.AutoPlay = True
    LoadResAVI aniPing, lResID

Exit Sub

bad_file:
    'Probably an error in the machines.dat file!
    MsgBox "There is an error in your MACHINES.DAT. Please correct it and try again!", 0 + 48, "Oops!"

End Sub

Private Sub Form_Load()

  Dim intLcount As Integer

    'Initialise timer
    tmrPing.Interval = 1000
    tmrPing.Enabled = False

    'Initialise spreadsheet
    With grdDetails
        .ColWidth(0) = 460
        .ColWidth(1) = 1700
        .ColWidth(2) = 1260
        .ColWidth(3) = 1000
        .ColWidth(4) = 2120
        .ColWidth(5) = 472
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 1
        .ColAlignment(4) = 1
        .ColAlignment(5) = 1
    End With 'grdDetails

    'Open the logging file
    fnum = FreeFile
    Open "c:\pinglog.txt" For Append As fnum
    Print #fnum, "Pinger Started: " & CStr(Now)
    Print #fnum, "==================================="
    Print #fnum,

    'Initialise the array
    For intLcount = 1 To 999
        ping_array(intLcount) = 0
    Next intLcount

    'Load the animation
    aniPing.AutoPlay = False
    lResID = 101
    LoadResAVI aniPing, lResID

End Sub

Private Sub Form_Resize()

  'Don't let them resize the form!

    If frmMain.WindowState <> 1 Then
        frmMain.Width = 7680
        frmMain.Height = 7170
    End If

End Sub

Private Sub grdDetails_Click()

  Dim intLogFile As Integer
  Dim intLcount As Integer
  Dim writestring As String

    'Set logging on/off picture
    With grdDetails
        If .Col = 5 Then
            If .CellPicture = Picture1(3).Picture Then
                Set .CellPicture = Picture1(2).Picture
              Else
                Set .CellPicture = Picture1(3).Picture
            End If
        End If
    End With 'grdDetails

    'Rewrite the machine.dat file to reflect changes
    intLogFile = FreeFile
    Open App.Path & "\machines.dat" For Output As intLogFile
    With grdDetails
        For intLcount = 1 To .Rows - 1
            .Row = intLcount
            .Col = 1
            writestring = .Text & ","
            .Col = 2
            writestring = writestring & .Text & ","
            .Col = 5

            'Show logging
            If .CellPicture = Picture1(2).Picture Then
                writestring = writestring & "1"
              Else
                writestring = writestring & "0"
            End If
            
            Print #intLogFile, writestring
        Next intLcount
    End With
    Close #intLogFile

End Sub

Private Sub tmrPing_Timer()

  'Increment the timer counter and look for 'ping time'

    timer_count = timer_count + 1
    If frmMain.WindowState = 1 Then
        frmMain.Caption = "Pinger: " & (delay * 60) - timer_count
      Else
        frmMain.Caption = "Pinger"
    End If

    'Show timer status
    staStatus.Panels(1).Text = "Ping Time: " & CStr((delay * 60) - timer_count)

    'If it's 'ping time', PING!!
    If timer_count >= (delay * 60) Then
        tmrPing.Enabled = False
        do_ping
        timer_count = 0
        tmrPing.Enabled = True
    End If

End Sub

Private Sub tbrTools_ButtonClick(ByVal Button As MSComctlLib.Button)

  Dim intLcount As Integer
  Dim intCheck As Integer
  Dim strExecute As String

    'Get the button clicks
    Select Case Button.Index

      Case 1
        'Force a ping now!
        tmrPing.Enabled = False
        DoEvents
        do_ping
        timer_count = 0
        tmrPing.Enabled = True

      Case 2
        'ExecCmd is a custom function that waits till the shelled
        'app has closed before continuing
        tmrPing.Enabled = False
        strExecute = "NOTEPAD " & App.Path & "\machines.dat"
        intCheck = ExecCmd(strExecute)
        Call Form_Activate

      Case 3
        'Load the configuration form
        tmrPing.Enabled = False
        frmConfig.Show

      Case 4
        'Quit the program
        Unload Me
        End

    End Select

End Sub
