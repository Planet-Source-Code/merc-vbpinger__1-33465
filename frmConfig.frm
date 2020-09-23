VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmConfig 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuration"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5460
   Icon            =   "frmConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComCtl2.UpDown udPoll 
      Height          =   285
      Left            =   4231
      TabIndex        =   4
      Top             =   450
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   503
      _Version        =   393216
      Value           =   1
      AutoBuddy       =   -1  'True
      BuddyControl    =   "lblPings"
      BuddyDispid     =   196613
      OrigLeft        =   4470
      OrigTop         =   450
      OrigRight       =   4710
      OrigBottom      =   735
      Max             =   100
      Min             =   1
      Wrap            =   -1  'True
      Enabled         =   -1  'True
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   2775
      TabIndex        =   1
      Top             =   1440
      Width           =   1395
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   390
      Left            =   1290
      TabIndex        =   0
      Top             =   1440
      Width           =   1395
   End
   Begin MSComCtl2.UpDown udReport 
      Height          =   285
      Left            =   4231
      TabIndex        =   7
      Top             =   840
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   503
      _Version        =   393216
      Value           =   1
      AutoBuddy       =   -1  'True
      BuddyControl    =   "lblDelay"
      BuddyDispid     =   196611
      OrigLeft        =   4470
      OrigTop         =   885
      OrigRight       =   4710
      OrigBottom      =   1170
      Max             =   100
      Min             =   1
      Wrap            =   -1  'True
      Enabled         =   -1  'True
   End
   Begin VB.Label lblDelay 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3855
      TabIndex        =   6
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "POLL Interval (mins)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   375
      TabIndex        =   5
      Top             =   847
      Width           =   3525
   End
   Begin VB.Label lblPings 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3855
      TabIndex        =   3
      Top             =   450
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Number of POLLs before reporting failure:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   375
      TabIndex        =   2
      Top             =   450
      Width           =   3525
   End
   Begin VB.Shape Shape1 
      Height          =   1785
      Left            =   75
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()

  'Save the new vals to the registry

    SetBinaryValue "HKEY_LOCAL_MACHINE\SOFTWARE\ICMP Pinger", "Pings", lblPings.Caption
    SetBinaryValue "HKEY_LOCAL_MACHINE\SOFTWARE\ICMP Pinger", "Delay", lblDelay.Caption

    'Set new times
    pings = lblPings.Caption
    delay = lblDelay.Caption

    'And unload the form
    Unload Me

End Sub

Private Sub cmdCancel_Click()

  'Unload the form

    Unload Me

End Sub

Private Sub Form_Load()

  'Load current settings and fill out form

    lblPings.Caption = GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\ICMP Pinger", "Pings")
    lblDelay.Caption = GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\ICMP Pinger", "Delay")
    udPoll.Value = lblPings
    udReport.Value = lblDelay

End Sub

Private Sub udPoll_Change()

  'Show value

    lblPings.Caption = udPoll.Value

End Sub

Private Sub udReport_Change()

  'Show value

    lblDelay.Caption = udReport.Value

End Sub
