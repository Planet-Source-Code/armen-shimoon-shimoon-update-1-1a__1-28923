VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmWizard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shimoon Update - v1.1a"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7620
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWizard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   7620
   StartUpPosition =   3  'Windows Default
   Begin InetCtlsObjects.Inet net 
      Left            =   6960
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.PictureBox frame32 
      BorderStyle     =   0  'None
      Height          =   5055
      Left            =   2520
      ScaleHeight     =   5055
      ScaleWidth      =   4935
      TabIndex        =   26
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
      Begin VB.CommandButton cmdFinish1 
         Caption         =   "&Finish"
         Height          =   375
         Left            =   3600
         TabIndex        =   29
         Top             =   4560
         Width           =   1215
      End
      Begin VB.Label Label16 
         Caption         =   "You have the latest version of the software. No updates are currently available."
         Height          =   735
         Left            =   120
         TabIndex        =   28
         Top             =   480
         Width           =   3495
      End
      Begin VB.Label Label15 
         Caption         =   "Step 3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.PictureBox frame42 
      BorderStyle     =   0  'None
      Height          =   5055
      Left            =   2520
      ScaleHeight     =   5055
      ScaleWidth      =   4935
      TabIndex        =   34
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
      Begin VB.CommandButton cmdFinish3 
         Caption         =   "&Finish"
         Height          =   375
         Left            =   3600
         TabIndex        =   37
         Top             =   4560
         Width           =   1215
      End
      Begin VB.Label Label20 
         Caption         =   "An update was downloaded but not installed. Click finish to close the wizard."
         Height          =   735
         Left            =   120
         TabIndex        =   36
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label Label19 
         Caption         =   "Step 4"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.PictureBox frame41 
      BorderStyle     =   0  'None
      Height          =   5055
      Left            =   2520
      ScaleHeight     =   5055
      ScaleWidth      =   4935
      TabIndex        =   30
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
      Begin VB.CommandButton cmdFinish2 
         Caption         =   "&Finish"
         Height          =   375
         Left            =   3600
         TabIndex        =   33
         Top             =   4560
         Width           =   1215
      End
      Begin VB.Label Label18 
         Caption         =   "Thank you. The update will begin once you close this wizard."
         Height          =   615
         Left            =   120
         TabIndex        =   32
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label Label17 
         Caption         =   "Step 4"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.PictureBox frame31 
      BorderStyle     =   0  'None
      Height          =   5055
      Left            =   2520
      ScaleHeight     =   5055
      ScaleWidth      =   4935
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   4935
      Begin VB.CommandButton cmdCancel3 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   2400
         TabIndex        =   25
         Top             =   4560
         Width           =   1095
      End
      Begin VB.CommandButton cmdNext3 
         Caption         =   "&Next ->"
         Height          =   375
         Left            =   3600
         TabIndex        =   24
         Top             =   4560
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "No"
         Height          =   375
         Left            =   960
         TabIndex        =   23
         Top             =   1440
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Yes"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   1440
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.Label Label14 
         Caption         =   "Would you like to install this update now?"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Width           =   4575
      End
      Begin VB.Label Label13 
         Caption         =   "The wizard has finished downloading the data. Your version is outdated, and the wizard has downloaded the update file. "
         Height          =   735
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label Label12 
         Caption         =   "Step 3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.PictureBox frame2 
      BorderStyle     =   0  'None
      Height          =   5055
      Left            =   2520
      ScaleHeight     =   5055
      ScaleWidth      =   4935
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
      Begin VB.CommandButton cmdStart 
         Caption         =   "&Start"
         Height          =   375
         Left            =   120
         TabIndex        =   38
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel2 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   2400
         TabIndex        =   17
         Top             =   4560
         Width           =   1095
      End
      Begin VB.CommandButton cmdNext2 
         Caption         =   "&Next ->"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3600
         TabIndex        =   16
         Top             =   4560
         Width           =   1215
      End
      Begin VB.Label lblStat 
         Alignment       =   2  'Center
         Caption         =   "Idle..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2280
         Width           =   4695
      End
      Begin VB.Label Label11 
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "The wizard is now connecting to the server, and downloading requested data."
         Height          =   615
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   4455
      End
      Begin VB.Label Label9 
         Caption         =   "Step 2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.PictureBox frame1 
      BorderStyle     =   0  'None
      Height          =   5055
      Left            =   2520
      ScaleHeight     =   5055
      ScaleWidth      =   4935
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      Begin VB.CommandButton cmdCancel1 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   2400
         TabIndex        =   10
         Top             =   4560
         Width           =   1095
      End
      Begin VB.CommandButton cmdNext1 
         Caption         =   "&Next ->"
         Height          =   375
         Left            =   3600
         TabIndex        =   9
         Top             =   4560
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "-If the user wants to download the update, it will download the update into a local file."
         Height          =   495
         Left            =   480
         TabIndex        =   8
         Top             =   2640
         Width           =   4095
      End
      Begin VB.Label Label7 
         Caption         =   "-A message to explain to the user about the update."
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   2400
         Width           =   3975
      End
      Begin VB.Label Label6 
         Caption         =   "-The URL for the new update file."
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   2160
         Width           =   3015
      End
      Begin VB.Label Label5 
         Caption         =   "-The new update version."
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   1920
         Width           =   3255
      End
      Begin VB.Label Label4 
         Caption         =   "This wizard will get the following information:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   4335
      End
      Begin VB.Label Label3 
         Caption         =   "Data to retrieve"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   $"frmWizard.frx":000C
         Height          =   735
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label Label1 
         Caption         =   "Step 1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Please vote for me on PSC if you liked this code or found it somewhat usefull. Please be nice and play fair!"
      Height          =   855
      Left            =   120
      TabIndex        =   39
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Please vote for me on PSC if you liked this code or found it somewhat usefull. Please be nice and play fair!"
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   130
      TabIndex        =   40
      Top             =   3740
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   5250
      Left            =   0
      Picture         =   "frmWizard.frx":00AA
      Top             =   0
      Width           =   2400
   End
End
Attribute VB_Name = "frmWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************
'Component: frmWizard Form
'Author: Armen Shimoon
'Copyright: Shimoon Technologies 2001
'Email: a_shimoon@hotmail.com
'****************************


'How to use? Refer to module for instructions

Private Sub cmdCancel1_Click()
Unload frmWizard
End
End Sub

Private Sub cmdCancel2_Click()
Unload frmWizard
End
End Sub

Private Sub cmdCancel3_Click()
Unload frmWizard
End
End Sub

Private Sub cmdFinish1_Click()

Unload frmWizard
End Sub

Private Sub cmdFinish2_Click()
Call Puttofile
End Sub

Private Sub cmdFinish3_Click()

Unload frmWizard

End Sub

Private Sub cmdNext1_Click()
frame1.Visible = False
frame2.Visible = True

End Sub

Private Sub cmdNext2_Click()
frame2.Visible = False

If nVer > Version Then
    frame31.Visible = True
Else
    frame32.Visible = True
End If
End Sub

Private Sub cmdNext3_Click()
If Option1.Value = True Then
    frame31.Visible = False
    frame41.Visible = True
Else
    frame31.Visible = False
    frame42.Visible = True
End If
    
End Sub

Private Sub cmdStart_Click()
Call GetData
End Sub

Private Sub Form_Load()

End Sub
