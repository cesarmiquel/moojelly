VERSION 5.00
Begin VB.Form Patching 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Patch Rom"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   6135
   Icon            =   "Patching.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   6135
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "Clear Log"
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   3360
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Patching.frx":014A
      Left            =   2160
      List            =   "Patching.frx":015D
      TabIndex        =   3
      Text            =   "All"
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Patch Level"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3720
      TabIndex        =   2
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Patch All"
      Height          =   255
      Left            =   2520
      TabIndex        =   1
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox status_text 
      Height          =   2655
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   600
      Width           =   5895
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Detail level:"
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Menu menu_file 
      Caption         =   "&File"
      Begin VB.Menu menu_export_log 
         Caption         =   "&Export Log"
      End
      Begin VB.Menu gap1 
         Caption         =   "-"
      End
      Begin VB.Menu menu_exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu menu_help 
      Caption         =   "&Help"
      Begin VB.Menu menu_howto 
         Caption         =   "&Howto Guide"
      End
      Begin VB.Menu menu_about 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Patching"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Combo1_Click()
Main.Refresh_Status True, False
End Sub

Private Sub Command1_Click()
Main.Patch_Rom
status_text.SelStart = Len(status_text)
End Sub

Private Sub Command3_Click()
status_text = vbNullString
Status_data = vbNullString
Main.Status "Log Cleared", 1, False, True
End Sub

Private Sub Form_Load()
status_text = ""
Combo1.ListIndex = 3
Main.Status "Log: Patching log started", 2
End Sub

Private Sub menu_about_Click()
About.Visible = True
End Sub

Private Sub menu_exit_Click()
Unload Me
End Sub

Private Sub menu_export_log_Click()
Main.menu_export_log_Click
End Sub

Private Sub menu_howto_Click()
On Error Resume Next
ShellExecute Me.hwnd, "open", App.Path & "\help\help.htm", "", "", 1
End Sub
