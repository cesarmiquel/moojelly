VERSION 5.00
Begin VB.Form About 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3975
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   249
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "SmellyMoo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   255
      Left            =   885
      TabIndex        =   15
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "CoolToby"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   600
      TabIndex        =   14
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Antonio Giuliana"
      Height          =   255
      Index           =   6
      Left            =   600
      TabIndex        =   13
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "JihadJoe"
      Height          =   255
      Index           =   5
      Left            =   2040
      TabIndex        =   12
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "RacoonSam"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   2040
      TabIndex        =   11
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Eastman"
      Height          =   255
      Index           =   3
      Left            =   2040
      TabIndex        =   10
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Thanks to these guys for Alpa testing or info:"
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   1560
      Width           =   3975
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Adrian"
      Height          =   255
      Index           =   2
      Left            =   600
      TabIndex        =   8
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "c0nfu53d"
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   7
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "MiniSamba"
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   6
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Don't hesitate to contact me @gmail.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   3000
      Width           =   3975
   End
   Begin VB.Label version 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Written by                     in 07-09 in homage to SML2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   3975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "'Super Mario Land 2/3' editor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "MooJelly"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
version.Caption = "version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Label5_Click()
SmellyMoo = True
Main.Frame9.Visible = True
End Sub
