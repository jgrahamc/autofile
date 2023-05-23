VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AutoFile "
   ClientHeight    =   9690
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   8385
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9690
   ScaleWidth      =   8385
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   6480
      TabIndex        =   5
      Top             =   9120
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Learning Progress"
      Height          =   6615
      Left            =   240
      TabIndex        =   4
      Top             =   2400
      Width           =   7815
      Begin MSComctlLib.TreeView tvwFolders 
         Height          =   4455
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   7858
         _Version        =   393217
         LabelEdit       =   1
         Sorted          =   -1  'True
         Style           =   7
         Appearance      =   1
      End
      Begin VB.Label lblStatus 
         Caption         =   "Ready"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   7335
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Learn"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "To teach AutoFile now about your Microsoft Outlook set up simply press Learn now."
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   8055
   End
   Begin VB.Label Label2 
      Caption         =   $"frmMain.frx":0442
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   7935
   End
   Begin VB.Label Label1 
      Caption         =   "Welcome to AutoFile"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7815
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
Private Sub Command1_Click()
    Command2.enabled = False
    Command1.enabled = False
    Call Learn(Me)
    Command2.enabled = True
End Sub
Private Sub Command2_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    Command1.enabled = True
    Call SaveAutoFileLocation
End Sub
Private Sub mnuHelpAbout_Click()
    frmAbout.Show 1
End Sub
