VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   2010
   ClientLeft      =   3495
   ClientTop       =   3435
   ClientWidth     =   6585
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1387.338
   ScaleMode       =   0  'User
   ScaleWidth      =   6183.655
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   240
      Picture         =   "frmAbout.frx":0442
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   5160
      TabIndex        =   0
      Top             =   120
      Width           =   1260
   End
   Begin VB.Label lblDescription 
      Caption         =   "autofile"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   1080
      TabIndex        =   2
      Top             =   1125
      Width           =   3870
   End
   Begin VB.Label lblTitle 
      Caption         =   "Application Title"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1080
      TabIndex        =   4
      Top             =   240
      Width           =   3870
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      Height          =   225
      Left            =   1080
      TabIndex        =   5
      Top             =   840
      Width           =   3870
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "Copyright (c) 2002 John Graham-Cumming"
      ForeColor       =   &H00000000&
      Height          =   465
      Left            =   1080
      TabIndex        =   3
      Top             =   1440
      Width           =   3870
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdOK_Click()
  Unload Me
End Sub
Private Sub Form_Load()
    Me.Caption = "About " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.FileDescription
    lblDisclaimer.Caption = App.LegalCopyright
End Sub

