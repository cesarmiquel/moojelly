VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Main 
   Appearance      =   0  'Flat
   Caption         =   "MooJelly"
   ClientHeight    =   8865
   ClientLeft      =   2070
   ClientTop       =   1410
   ClientWidth     =   10050
   FillStyle       =   0  'Solid
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8865
   ScaleWidth      =   10050
   Begin VB.PictureBox temp 
      Height          =   975
      Left            =   7680
      ScaleHeight     =   61
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   117
      Top             =   6120
      Visible         =   0   'False
      Width           =   1575
      Begin VB.Timer reg_check 
         Interval        =   3000
         Left            =   480
         Top             =   360
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   120
         Top             =   360
      End
      Begin VB.PictureBox sprites1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0000FF00&
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   120
         ScaleHeight     =   7
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   7
         TabIndex        =   119
         Top             =   120
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.PictureBox special_blocks1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0000FF00&
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   360
         ScaleHeight     =   7
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   7
         TabIndex        =   118
         Top             =   120
         Visible         =   0   'False
         Width           =   135
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   960
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   8775
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   15478
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Level"
      TabPicture(0)   =   "main.frx":014A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label11"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "SSTab1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Combo2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "combo1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Global"
      TabPicture(1)   =   "main.frx":0166
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SSTab3"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Config"
      TabPicture(2)   =   "main.frx":0182
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).Control(1)=   "Frame1"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Status"
      TabPicture(3)   =   "main.frx":019E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Combo4"
      Tab(3).Control(1)=   "status_text"
      Tab(3).Control(2)=   "Command3"
      Tab(3).Control(3)=   "Label18"
      Tab(3).ControlCount=   4
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "main.frx":01BA
         Left            =   -73680
         List            =   "main.frx":01CD
         TabIndex        =   76
         Text            =   "All"
         Top             =   720
         Width           =   1695
      End
      Begin VB.Frame Frame2 
         Caption         =   "Function Config"
         Height          =   2175
         Left            =   -71880
         TabIndex        =   50
         Top             =   600
         Width           =   2655
         Begin VB.CheckBox Check2 
            Caption         =   "Load config from .ini"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   55
            Top             =   600
            Value           =   1  'Checked
            Width           =   2175
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Use Tile set"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   54
            Top             =   1560
            Width           =   1695
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Use Block data"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   53
            Top             =   1320
            Width           =   1935
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Use headers"
            Enabled         =   0   'False
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   52
            Top             =   1080
            Width           =   1695
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Load unsupported ROMs"
            Enabled         =   0   'False
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   51
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Graphics"
         Height          =   2175
         Left            =   -74760
         TabIndex        =   45
         Top             =   600
         Width           =   2655
         Begin VB.CheckBox Check5 
            Caption         =   "Colorize"
            Enabled         =   0   'False
            Height          =   255
            Left            =   240
            TabIndex        =   114
            Top             =   1320
            Value           =   1  'Checked
            Width           =   1815
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Display Sprites"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   49
            Top             =   1680
            Value           =   1  'Checked
            Width           =   1575
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Type guess mode"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   48
            Top             =   1080
            Width           =   1575
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Smooth shrink (slower)"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   47
            Top             =   360
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Smart (hidden blocks)"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   46
            Top             =   840
            Value           =   1  'Checked
            Width           =   2295
         End
      End
      Begin VB.TextBox status_text 
         Height          =   6975
         Left            =   -74760
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   44
         Top             =   1200
         Width           =   9615
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Clear Log"
         Height          =   255
         Left            =   -74640
         TabIndex        =   43
         Top             =   8280
         Width           =   855
      End
      Begin TabDlg.SSTab SSTab3 
         Height          =   8175
         Left            =   -75000
         TabIndex        =   1
         Top             =   600
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   14420
         _Version        =   393216
         Tabs            =   4
         Tab             =   3
         TabsPerRow      =   5
         TabHeight       =   520
         TabCaption(0)   =   "Edit Tile-Map"
         TabPicture(0)   =   "main.frx":0201
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Label9(1)"
         Tab(0).Control(1)=   "Label10(1)"
         Tab(0).Control(2)=   "Label9(2)"
         Tab(0).Control(3)=   "Combo3"
         Tab(0).Control(4)=   "Picture4"
         Tab(0).Control(5)=   "tilemap_VScroll"
         Tab(0).Control(6)=   "tilemap_HScroll"
         Tab(0).Control(7)=   "Frame6"
         Tab(0).Control(8)=   "Check9(0)"
         Tab(0).Control(9)=   "Check9(1)"
         Tab(0).Control(10)=   "Check9(2)"
         Tab(0).Control(11)=   "ToolBox2(21)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "ToolBox2(20)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "ToolBox2(19)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "ToolBox2(18)"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "ToolBox2(15)"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "ToolBox2(14)"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "Command5"
         Tab(0).Control(18)=   "Text4"
         Tab(0).ControlCount=   19
         TabCaption(1)   =   "Credits text"
         TabPicture(1)   =   "main.frx":021D
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Command2"
         Tab(1).Control(1)=   "Command1"
         Tab(1).Control(2)=   "credits_Text1"
         Tab(1).Control(3)=   "Command4"
         Tab(1).ControlCount=   4
         TabCaption(2)   =   "Credits graphics"
         TabPicture(2)   =   "main.frx":0239
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "credits_Picture1"
         Tab(2).Control(1)=   "credits_VScroll1"
         Tab(2).Control(2)=   "Frame7"
         Tab(2).ControlCount=   3
         TabCaption(3)   =   "Music"
         TabPicture(3)   =   "main.frx":0255
         Tab(3).ControlEnabled=   -1  'True
         Tab(3).Control(0)=   "Label20"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).ControlCount=   1
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   -74640
            TabIndex        =   141
            Text            =   "1"
            Top             =   5160
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Command5"
            Height          =   255
            Left            =   -74640
            TabIndex        =   140
            Top             =   4680
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.CommandButton ToolBox2 
            Height          =   375
            Index           =   14
            Left            =   -74880
            Picture         =   "main.frx":0271
            Style           =   1  'Graphical
            TabIndex        =   58
            TabStop         =   0   'False
            ToolTipText     =   "Block Pen"
            Top             =   1320
            Width           =   375
         End
         Begin VB.CommandButton ToolBox2 
            Height          =   375
            Index           =   15
            Left            =   -74520
            Picture         =   "main.frx":095B
            Style           =   1  'Graphical
            TabIndex        =   59
            TabStop         =   0   'False
            ToolTipText     =   "Sample Block Type"
            Top             =   1320
            Width           =   375
         End
         Begin VB.CommandButton ToolBox2 
            Enabled         =   0   'False
            Height          =   375
            Index           =   18
            Left            =   -74880
            Picture         =   "main.frx":1045
            Style           =   1  'Graphical
            TabIndex        =   60
            TabStop         =   0   'False
            ToolTipText     =   "Block Fill"
            Top             =   1680
            Width           =   375
         End
         Begin VB.CommandButton ToolBox2 
            Enabled         =   0   'False
            Height          =   375
            Index           =   19
            Left            =   -74520
            Picture         =   "main.frx":172F
            Style           =   1  'Graphical
            TabIndex        =   61
            TabStop         =   0   'False
            ToolTipText     =   "Sample Block Type"
            Top             =   1680
            Width           =   375
         End
         Begin VB.CommandButton ToolBox2 
            Enabled         =   0   'False
            Height          =   375
            Index           =   20
            Left            =   -74880
            Picture         =   "main.frx":1E19
            Style           =   1  'Graphical
            TabIndex        =   62
            TabStop         =   0   'False
            ToolTipText     =   "Magnifier (zoom in)"
            Top             =   2040
            Width           =   375
         End
         Begin VB.CommandButton ToolBox2 
            Enabled         =   0   'False
            Height          =   375
            Index           =   21
            Left            =   -74520
            Picture         =   "main.frx":2503
            Style           =   1  'Graphical
            TabIndex        =   63
            TabStop         =   0   'False
            ToolTipText     =   "Magnifier (zoom out)"
            Top             =   2040
            Width           =   375
         End
         Begin VB.CheckBox Check9 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Paths"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   2
            Left            =   -74880
            Style           =   1  'Graphical
            TabIndex        =   135
            Top             =   4080
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.CheckBox Check9 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Colorise"
            Enabled         =   0   'False
            ForeColor       =   &H0000FF00&
            Height          =   375
            Index           =   1
            Left            =   -74880
            Style           =   1  'Graphical
            TabIndex        =   134
            Top             =   3720
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.CheckBox Check9 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Tiles"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   0
            Left            =   -74880
            Style           =   1  'Graphical
            TabIndex        =   133
            Top             =   3360
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.Frame Frame7 
            Caption         =   "Tools"
            Height          =   3615
            Left            =   -71640
            TabIndex        =   86
            Top             =   960
            Width           =   3255
            Begin VB.PictureBox credits_tiles1 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H0000FF00&
               ForeColor       =   &H80000008&
               Height          =   1945
               Left            =   240
               ScaleHeight     =   128
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   128
               TabIndex        =   88
               Top             =   720
               Width           =   1945
               Begin VB.Shape Shape6 
                  BorderColor     =   &H00000000&
                  DrawMode        =   10  'Mask Pen
                  Height          =   120
                  Left            =   0
                  Shape           =   1  'Square
                  Top             =   0
                  Width           =   120
               End
            End
            Begin VB.CheckBox Check4 
               Caption         =   "Display screen border lines"
               Height          =   375
               Left            =   240
               TabIndex        =   87
               Top             =   3000
               Width           =   2295
            End
            Begin VB.Label Label15 
               Caption         =   "Palette:"
               Height          =   255
               Left            =   2520
               TabIndex        =   94
               Top             =   360
               Width           =   615
            End
            Begin VB.Label palette3 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   2520
               TabIndex        =   93
               Top             =   660
               Width           =   255
            End
            Begin VB.Label palette3 
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   1
               Left            =   2520
               TabIndex        =   92
               Top             =   900
               Width           =   255
            End
            Begin VB.Label palette3 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   2
               Left            =   2520
               TabIndex        =   91
               Top             =   1140
               Width           =   255
            End
            Begin VB.Label palette3 
               Appearance      =   0  'Flat
               BackColor       =   &H000080FF&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   3
               Left            =   2520
               TabIndex        =   90
               Top             =   1380
               Width           =   255
            End
            Begin VB.Label Label17 
               Caption         =   "Tiles:"
               Height          =   255
               Left            =   240
               TabIndex        =   89
               Top             =   360
               Width           =   495
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Tools"
            Height          =   4095
            Left            =   -68040
            TabIndex        =   78
            Top             =   960
            Width           =   2415
            Begin VB.PictureBox tiles2 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H0000FF00&
               ForeColor       =   &H80000008&
               Height          =   1945
               Left            =   240
               ScaleHeight     =   128
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   128
               TabIndex        =   79
               Top             =   600
               Width           =   1945
               Begin VB.Shape Shape7 
                  BorderColor     =   &H00000000&
                  DrawMode        =   10  'Mask Pen
                  Height          =   120
                  Left            =   0
                  Top             =   0
                  Width           =   120
               End
            End
            Begin VB.Label Label1 
               Caption         =   "Tiles:"
               Height          =   255
               Left            =   240
               TabIndex        =   85
               Top             =   360
               Width           =   495
            End
            Begin VB.Label palette2 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   3
               Left            =   240
               TabIndex        =   84
               Top             =   3660
               Width           =   255
            End
            Begin VB.Label palette2 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   2
               Left            =   240
               TabIndex        =   83
               Top             =   3420
               Width           =   255
            End
            Begin VB.Label palette2 
               Appearance      =   0  'Flat
               BackColor       =   &H00008000&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   82
               Top             =   3180
               Width           =   255
            End
            Begin VB.Label palette2 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   81
               Top             =   2940
               Width           =   255
            End
            Begin VB.Label Label16 
               Caption         =   "Palette:"
               Height          =   255
               Left            =   240
               TabIndex        =   80
               Top             =   2640
               Width           =   735
            End
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Save"
            Enabled         =   0   'False
            Height          =   255
            Left            =   -71400
            TabIndex        =   75
            Top             =   7560
            Width           =   735
         End
         Begin VB.HScrollBar tilemap_HScroll 
            Height          =   255
            LargeChange     =   8
            Left            =   -73920
            Max             =   0
            TabIndex        =   74
            Top             =   7680
            Visible         =   0   'False
            Width           =   5415
         End
         Begin VB.VScrollBar tilemap_VScroll 
            Height          =   6735
            LargeChange     =   8
            Left            =   -68510
            Max             =   0
            TabIndex        =   73
            Top             =   960
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.VScrollBar credits_VScroll1 
            Height          =   6375
            LargeChange     =   32
            Left            =   -72090
            Max             =   3000
            Min             =   1
            SmallChange     =   8
            TabIndex        =   72
            Top             =   960
            Value           =   1
            Width           =   255
         End
         Begin VB.PictureBox credits_Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            ForeColor       =   &H80000008&
            Height          =   6375
            Left            =   -74520
            ScaleHeight     =   423
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   160
            TabIndex        =   70
            Top             =   960
            Width           =   2435
            Begin VB.PictureBox credits1 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   150
               Left            =   -15
               ScaleHeight     =   10
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   162
               TabIndex        =   71
               Top             =   -15
               Width           =   2430
            End
         End
         Begin VB.TextBox credits_Text1 
            Height          =   6495
            Left            =   -74040
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   69
            Top             =   960
            Width           =   4335
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Revert"
            Height          =   255
            Left            =   -73080
            TabIndex        =   68
            Top             =   7560
            Width           =   735
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Refresh"
            Enabled         =   0   'False
            Height          =   255
            Left            =   -72240
            TabIndex        =   67
            Top             =   7560
            Width           =   735
         End
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00808080&
            ForeColor       =   &H80000008&
            Height          =   6735
            Left            =   -73920
            ScaleHeight     =   447
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   359
            TabIndex        =   65
            Top             =   960
            Width           =   5415
            Begin VB.PictureBox tilemap2 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H0000FF00&
               ForeColor       =   &H80000008&
               Height          =   3795
               Left            =   -15
               ScaleHeight     =   251
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   250
               TabIndex        =   66
               Top             =   -15
               Width           =   3780
            End
         End
         Begin VB.ComboBox Combo3 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "main.frx":2BED
            Left            =   -74040
            List            =   "main.frx":2BEF
            TabIndex        =   64
            Text            =   "Map"
            Top             =   480
            Width           =   2655
         End
         Begin VB.Label Label20 
            Caption         =   "Work in progress, features disabled in this version"
            Height          =   255
            Left            =   1320
            TabIndex        =   142
            Top             =   1680
            Width           =   3615
         End
         Begin VB.Label Label9 
            Caption         =   "Display:"
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
            Index           =   2
            Left            =   -74880
            TabIndex        =   132
            Top             =   3120
            Width           =   735
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            Caption         =   "R"
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
            Index           =   1
            Left            =   -74520
            TabIndex        =   57
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            Caption         =   "L"
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
            Index           =   1
            Left            =   -74880
            TabIndex        =   56
            Top             =   1080
            Width           =   375
         End
      End
      Begin VB.ComboBox combo1 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "main.frx":2BF1
         Left            =   240
         List            =   "main.frx":2BF3
         TabIndex        =   4
         Text            =   "Zone"
         Top             =   480
         Width           =   1575
      End
      Begin VB.ComboBox Combo2 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1920
         TabIndex        =   3
         Text            =   "Level"
         Top             =   480
         Width           =   2415
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   7815
         Left            =   0
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   960
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   13785
         _Version        =   393216
         Tabs            =   4
         TabsPerRow      =   5
         TabHeight       =   520
         TabCaption(0)   =   "Edit Level"
         TabPicture(0)   =   "main.frx":2BF5
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label10(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label9(0)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label7"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label19"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "ToolBox1(15)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "ToolBox1(14)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "VScroll3"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Picture3"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "ToolBox1(13)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "ToolBox1(12)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "ToolBox1(11)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "ToolBox1(10)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "ToolBox1(9)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "ToolBox1(8)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "MiniMap1"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "ToolBox1(7)"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "ToolBox1(6)"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "ToolBox1(5)"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "ToolBox1(4)"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "ToolBox1(3)"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "ToolBox1(2)"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "Pattern1"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "Blocks2"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).Control(23)=   "ToolBox1(1)"
         Tab(0).Control(23).Enabled=   0   'False
         Tab(0).Control(24)=   "ToolBox1(0)"
         Tab(0).Control(24).Enabled=   0   'False
         Tab(0).Control(25)=   "VScroll1"
         Tab(0).Control(25).Enabled=   0   'False
         Tab(0).Control(26)=   "HScroll1"
         Tab(0).Control(26).Enabled=   0   'False
         Tab(0).Control(27)=   "Picture2"
         Tab(0).Control(27).Enabled=   0   'False
         Tab(0).Control(28)=   "Check8(0)"
         Tab(0).Control(28).Enabled=   0   'False
         Tab(0).Control(29)=   "Check8(1)"
         Tab(0).Control(29).Enabled=   0   'False
         Tab(0).Control(30)=   "Check8(2)"
         Tab(0).Control(30).Enabled=   0   'False
         Tab(0).Control(31)=   "Check8(3)"
         Tab(0).Control(31).Enabled=   0   'False
         Tab(0).ControlCount=   32
         TabCaption(1)   =   "Edit Blocks"
         TabPicture(1)   =   "main.frx":2C11
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "palette1(3)"
         Tab(1).Control(1)=   "palette1(2)"
         Tab(1).Control(2)=   "palette1(1)"
         Tab(1).Control(3)=   "palette1(0)"
         Tab(1).Control(4)=   "Label6"
         Tab(1).Control(5)=   "Label3"
         Tab(1).Control(6)=   "Label2"
         Tab(1).Control(7)=   "Blocks1"
         Tab(1).Control(8)=   "Tiles1"
         Tab(1).Control(9)=   "Frame4"
         Tab(1).ControlCount=   10
         TabCaption(2)   =   "View Sprites"
         TabPicture(2)   =   "main.frx":2C2D
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label8"
         Tab(2).Control(1)=   "VScroll2"
         Tab(2).Control(2)=   "HScroll2"
         Tab(2).Control(3)=   "Picture1"
         Tab(2).ControlCount=   4
         TabCaption(3)   =   "Advanced"
         TabPicture(3)   =   "main.frx":2C49
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Frame5"
         Tab(3).Control(1)=   "Frame3"
         Tab(3).Control(2)=   "Frame8"
         Tab(3).Control(3)=   "Frame9"
         Tab(3).ControlCount=   4
         Begin VB.Frame Frame9 
            Caption         =   "smellymoo"
            Height          =   2535
            Left            =   -71520
            TabIndex        =   130
            Top             =   3480
            Visible         =   0   'False
            Width           =   3375
            Begin VB.TextBox Text3 
               Height          =   285
               Left            =   1320
               TabIndex        =   138
               Top             =   360
               Width           =   495
            End
            Begin VB.Label Label14 
               Caption         =   "Rom Version:"
               Height          =   255
               Index           =   2
               Left            =   240
               TabIndex        =   139
               Top             =   360
               Width           =   975
            End
         End
         Begin VB.CheckBox Check8 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Borders"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   3
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   129
            Top             =   5280
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.CheckBox Check8 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Minimap"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   2
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   128
            Top             =   4920
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.CheckBox Check8 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            DownPicture     =   "main.frx":2C65
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   1
            Left            =   120
            Picture         =   "main.frx":334F
            Style           =   1  'Graphical
            TabIndex        =   127
            Top             =   4560
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.CheckBox Check8 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Colorise"
            Enabled         =   0   'False
            ForeColor       =   &H0000FF00&
            Height          =   375
            Index           =   0
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   126
            Top             =   4200
            UseMaskColor    =   -1  'True
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.Frame Frame8 
            Caption         =   "Global Settings"
            Height          =   2295
            Left            =   -71520
            TabIndex        =   120
            Top             =   840
            Width           =   3375
            Begin VB.TextBox Text2 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1440
               TabIndex        =   122
               Top             =   480
               Width           =   1575
            End
            Begin VB.Label Label14 
               Caption         =   "Rom Name:"
               Height          =   255
               Index           =   0
               Left            =   480
               TabIndex        =   121
               Top             =   480
               Width           =   975
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Settings"
            Height          =   1935
            Left            =   -71400
            TabIndex        =   110
            Top             =   1200
            Width           =   2295
            Begin VB.CheckBox Check7 
               Caption         =   "Colorize"
               Enabled         =   0   'False
               Height          =   255
               Left            =   120
               TabIndex        =   116
               Top             =   1560
               Value           =   1  'Checked
               Width           =   1815
            End
            Begin VB.CheckBox Check6 
               Caption         =   "Type guess mode"
               Enabled         =   0   'False
               Height          =   255
               Left            =   120
               TabIndex        =   115
               Top             =   1320
               Width           =   1935
            End
            Begin VB.CheckBox Check3 
               Caption         =   "Use Tile set"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   113
               Top             =   600
               Width           =   1695
            End
            Begin VB.CheckBox Check3 
               Caption         =   "Use Block data"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   112
               Top             =   360
               Width           =   1935
            End
            Begin VB.CheckBox Check3 
               Caption         =   "Smart (hidden blocks)"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   111
               Top             =   1080
               Value           =   1  'Checked
               Width           =   1935
            End
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            ForeColor       =   &H80000008&
            Height          =   2895
            Left            =   -74520
            ScaleHeight     =   191
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   199
            TabIndex        =   108
            Top             =   1320
            Width           =   3015
            Begin VB.PictureBox sprites2 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H0000FF00&
               ForeColor       =   &H80000008&
               Height          =   7710
               Left            =   -15
               ScaleHeight     =   512
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   256
               TabIndex        =   109
               Top             =   -15
               Width           =   3870
            End
         End
         Begin VB.HScrollBar HScroll2 
            Height          =   255
            LargeChange     =   32
            Left            =   -74520
            Max             =   512
            SmallChange     =   16
            TabIndex        =   106
            Top             =   4200
            Width           =   3015
         End
         Begin VB.VScrollBar VScroll2 
            Height          =   2895
            LargeChange     =   32
            Left            =   -71520
            Max             =   512
            SmallChange     =   16
            TabIndex        =   105
            Top             =   1320
            Width           =   255
         End
         Begin VB.Frame Frame3 
            Caption         =   "Level Settings"
            Height          =   2295
            Left            =   -74400
            TabIndex        =   100
            Top             =   840
            Width           =   2655
            Begin VB.CheckBox Check10 
               Alignment       =   1  'Right Justify
               Caption         =   "Scrolling"
               Enabled         =   0   'False
               Height          =   255
               Left            =   600
               TabIndex        =   136
               Top             =   1080
               Width           =   975
            End
            Begin VB.TextBox Text1 
               Enabled         =   0   'False
               Height          =   285
               Index           =   0
               Left            =   1320
               Locked          =   -1  'True
               TabIndex        =   102
               Top             =   360
               Width           =   975
            End
            Begin VB.TextBox Text1 
               Enabled         =   0   'False
               Height          =   285
               Index           =   1
               Left            =   1320
               Locked          =   -1  'True
               TabIndex        =   101
               Top             =   720
               Width           =   975
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Music:"
               Height          =   255
               Left            =   240
               TabIndex        =   104
               Top             =   360
               Width           =   975
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Time Limit:"
               Height          =   255
               Left            =   240
               TabIndex        =   103
               Top             =   720
               Width           =   975
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Advanced Settings"
            Height          =   2535
            Left            =   -74520
            TabIndex        =   95
            Top             =   3480
            Width           =   2775
            Begin VB.TextBox Text1 
               Enabled         =   0   'False
               Height          =   285
               Index           =   4
               Left            =   1320
               Locked          =   -1  'True
               TabIndex        =   123
               Top             =   1080
               Width           =   975
            End
            Begin VB.TextBox Text1 
               Enabled         =   0   'False
               Height          =   285
               Index           =   2
               Left            =   1320
               Locked          =   -1  'True
               TabIndex        =   97
               Top             =   360
               Width           =   975
            End
            Begin VB.TextBox Text1 
               Enabled         =   0   'False
               Height          =   285
               Index           =   3
               Left            =   1320
               Locked          =   -1  'True
               TabIndex        =   96
               Top             =   720
               Width           =   975
            End
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
               Caption         =   "Block Data:"
               Height          =   255
               Index           =   1
               Left            =   360
               TabIndex        =   124
               Top             =   1080
               Width           =   855
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               Caption         =   "Map Bank:"
               Height          =   255
               Left            =   360
               TabIndex        =   99
               Top             =   360
               Width           =   855
            End
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
               Caption         =   "Sub Bank:"
               Height          =   255
               Index           =   0
               Left            =   360
               TabIndex        =   98
               Top             =   720
               Width           =   855
            End
         End
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00808080&
            Height          =   3255
            Left            =   960
            ScaleHeight     =   213
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   493
            TabIndex        =   31
            Top             =   1440
            Width           =   7455
            Begin VB.PictureBox TileMap1 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00808080&
               FillColor       =   &H000080FF&
               ForeColor       =   &H000000FF&
               Height          =   3495
               Left            =   -15
               ScaleHeight     =   231
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   272
               TabIndex        =   32
               Top             =   -15
               Width           =   4110
               Begin VB.Shape Shape3 
                  BorderColor     =   &H00000000&
                  DrawMode        =   10  'Mask Pen
                  Height          =   240
                  Left            =   0
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   240
               End
            End
         End
         Begin VB.HScrollBar HScroll1 
            Height          =   255
            LargeChange     =   320
            Left            =   960
            TabIndex        =   30
            Top             =   4680
            Width           =   7455
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   3315
            LargeChange     =   160
            Left            =   8400
            TabIndex        =   29
            Top             =   1380
            Width           =   255
         End
         Begin VB.PictureBox Tiles1 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H0000FF00&
            ForeColor       =   &H80000008&
            Height          =   1945
            Left            =   -74520
            ScaleHeight     =   128
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   128
            TabIndex        =   28
            Top             =   1260
            Width           =   1945
            Begin VB.Shape Shape2 
               BorderColor     =   &H00000000&
               DrawMode        =   10  'Mask Pen
               Height          =   120
               Left            =   0
               Shape           =   1  'Square
               Top             =   0
               Width           =   120
            End
         End
         Begin VB.PictureBox Blocks1 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H0000FF00&
            ForeColor       =   &H80000008&
            Height          =   1945
            Left            =   -74520
            ScaleHeight     =   128
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   256
            TabIndex        =   27
            Top             =   3780
            Width           =   3865
         End
         Begin VB.CommandButton ToolBox1 
            Height          =   375
            Index           =   0
            Left            =   120
            Picture         =   "main.frx":3A39
            Style           =   1  'Graphical
            TabIndex        =   26
            TabStop         =   0   'False
            ToolTipText     =   "Block Pen"
            Top             =   720
            Width           =   375
         End
         Begin VB.CommandButton ToolBox1 
            Height          =   375
            Index           =   1
            Left            =   480
            Picture         =   "main.frx":4123
            Style           =   1  'Graphical
            TabIndex        =   25
            TabStop         =   0   'False
            ToolTipText     =   "Sample Block Type"
            Top             =   720
            Width           =   375
         End
         Begin VB.PictureBox Blocks2 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00808080&
            ForeColor       =   &H80000008&
            Height          =   1945
            Left            =   960
            ScaleHeight     =   128
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   256
            TabIndex        =   24
            Top             =   5040
            Visible         =   0   'False
            Width           =   3865
            Begin VB.Shape Shape1 
               BorderColor     =   &H00000000&
               DrawMode        =   10  'Mask Pen
               Height          =   240
               Left            =   0
               Shape           =   1  'Square
               Top             =   0
               Width           =   240
            End
         End
         Begin VB.PictureBox Pattern1 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00808080&
            ForeColor       =   &H8000000E&
            Height          =   1935
            Left            =   5040
            ScaleHeight     =   127
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   159
            TabIndex        =   23
            Top             =   5040
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.CommandButton ToolBox1 
            Height          =   375
            Index           =   2
            Left            =   120
            Picture         =   "main.frx":480D
            Style           =   1  'Graphical
            TabIndex        =   22
            TabStop         =   0   'False
            ToolTipText     =   "Pattern Brush"
            Top             =   1080
            Width           =   375
         End
         Begin VB.CommandButton ToolBox1 
            Height          =   375
            Index           =   3
            Left            =   480
            Picture         =   "main.frx":4EF7
            Style           =   1  'Graphical
            TabIndex        =   21
            TabStop         =   0   'False
            ToolTipText     =   "Copy to Pattern Brush"
            Top             =   1080
            Width           =   375
         End
         Begin VB.CommandButton ToolBox1 
            Height          =   375
            Index           =   4
            Left            =   120
            Picture         =   "main.frx":55E1
            Style           =   1  'Graphical
            TabIndex        =   20
            TabStop         =   0   'False
            ToolTipText     =   "Block Fill"
            Top             =   1440
            Width           =   375
         End
         Begin VB.CommandButton ToolBox1 
            Height          =   375
            Index           =   5
            Left            =   480
            Picture         =   "main.frx":5CCB
            Style           =   1  'Graphical
            TabIndex        =   19
            TabStop         =   0   'False
            ToolTipText     =   "Sample Block Type"
            Top             =   1440
            Width           =   375
         End
         Begin VB.CommandButton ToolBox1 
            Height          =   375
            Index           =   6
            Left            =   120
            Picture         =   "main.frx":63B5
            Style           =   1  'Graphical
            TabIndex        =   18
            TabStop         =   0   'False
            ToolTipText     =   "Fill Line"
            Top             =   1800
            Width           =   375
         End
         Begin VB.CommandButton ToolBox1 
            Height          =   375
            Index           =   7
            Left            =   480
            Picture         =   "main.frx":6A9F
            Style           =   1  'Graphical
            TabIndex        =   17
            TabStop         =   0   'False
            ToolTipText     =   "Sample Block Type"
            Top             =   1800
            Width           =   375
         End
         Begin VB.PictureBox MiniMap1 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   715
            Left            =   960
            ScaleHeight     =   48
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   256
            TabIndex        =   16
            Top             =   480
            Visible         =   0   'False
            Width           =   3845
            Begin VB.Shape Shape4 
               BorderColor     =   &H000000FF&
               Height          =   720
               Left            =   2040
               Top             =   0
               Width           =   615
            End
         End
         Begin VB.CommandButton ToolBox1 
            Height          =   375
            Index           =   8
            Left            =   120
            Picture         =   "main.frx":7189
            Style           =   1  'Graphical
            TabIndex        =   15
            TabStop         =   0   'False
            ToolTipText     =   "History Brush"
            Top             =   2160
            Width           =   375
         End
         Begin VB.CommandButton ToolBox1 
            Height          =   375
            Index           =   9
            Left            =   480
            Picture         =   "main.frx":7873
            Style           =   1  'Graphical
            TabIndex        =   14
            TabStop         =   0   'False
            ToolTipText     =   "Magnifier"
            Top             =   2160
            Width           =   375
         End
         Begin VB.CommandButton ToolBox1 
            Height          =   375
            Index           =   10
            Left            =   120
            Picture         =   "main.frx":7F5D
            Style           =   1  'Graphical
            TabIndex        =   13
            TabStop         =   0   'False
            ToolTipText     =   "Sprite Tool"
            Top             =   2520
            Width           =   375
         End
         Begin VB.CommandButton ToolBox1 
            Height          =   375
            Index           =   11
            Left            =   480
            Picture         =   "main.frx":8647
            Style           =   1  'Graphical
            TabIndex        =   12
            TabStop         =   0   'False
            ToolTipText     =   "Mario Tool"
            Top             =   2520
            Width           =   375
         End
         Begin VB.CommandButton ToolBox1 
            Height          =   375
            Index           =   12
            Left            =   120
            Picture         =   "main.frx":8D31
            Style           =   1  'Graphical
            TabIndex        =   11
            TabStop         =   0   'False
            ToolTipText     =   "Magnifier (zoom in)"
            Top             =   2880
            Width           =   375
         End
         Begin VB.CommandButton ToolBox1 
            Height          =   375
            Index           =   13
            Left            =   480
            Picture         =   "main.frx":941B
            Style           =   1  'Graphical
            TabIndex        =   10
            TabStop         =   0   'False
            ToolTipText     =   "Magnifier (zoom out)"
            Top             =   2880
            Width           =   375
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00808080&
            ForeColor       =   &H80000008&
            Height          =   1945
            Left            =   960
            ScaleHeight     =   128
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   256
            TabIndex        =   7
            Top             =   5040
            Visible         =   0   'False
            Width           =   3865
            Begin VB.PictureBox sprites3 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H0000FF00&
               ForeColor       =   &H80000008&
               Height          =   7710
               Left            =   -15
               ScaleHeight     =   512
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   256
               TabIndex        =   9
               Top             =   -15
               Width           =   3870
               Begin VB.Shape Shape5 
                  BorderColor     =   &H000000FF&
                  Height          =   480
                  Left            =   480
                  Top             =   0
                  Width           =   480
               End
            End
         End
         Begin VB.VScrollBar VScroll3 
            Height          =   1945
            LargeChange     =   32
            Left            =   4830
            Max             =   256
            Min             =   1
            SmallChange     =   16
            TabIndex        =   5
            Top             =   5040
            Value           =   1
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CommandButton ToolBox1 
            Height          =   375
            Index           =   14
            Left            =   120
            Picture         =   "main.frx":9B05
            Style           =   1  'Graphical
            TabIndex        =   8
            TabStop         =   0   'False
            ToolTipText     =   "Border Tool"
            Top             =   3240
            Width           =   375
         End
         Begin VB.CommandButton ToolBox1 
            Caption         =   "X"
            Height          =   375
            Index           =   15
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   6
            TabStop         =   0   'False
            ToolTipText     =   "Events Tool (Advanced)"
            Top             =   3240
            Width           =   375
         End
         Begin VB.Label Label19 
            Caption         =   "Display:"
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
            Left            =   120
            TabIndex        =   125
            Top             =   3960
            Width           =   735
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Sprites:"
            Height          =   255
            Left            =   -74520
            TabIndex        =   107
            Top             =   1080
            Width           =   1935
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Tile Set:"
            Height          =   255
            Left            =   -74520
            TabIndex        =   42
            Top             =   1020
            Width           =   1935
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Block Data:"
            Height          =   255
            Left            =   -74520
            TabIndex        =   41
            Top             =   3540
            Width           =   3855
         End
         Begin VB.Label Label6 
            Caption         =   "Palette:"
            Height          =   255
            Left            =   -72360
            TabIndex        =   40
            Top             =   1020
            Width           =   735
         End
         Begin VB.Label palette1 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   -72360
            TabIndex        =   39
            Top             =   1320
            Width           =   255
         End
         Begin VB.Label palette1 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   -72360
            TabIndex        =   38
            Top             =   1560
            Width           =   255
         End
         Begin VB.Label palette1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   -72360
            TabIndex        =   37
            Top             =   1800
            Width           =   255
         End
         Begin VB.Label palette1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   -72360
            TabIndex        =   36
            Top             =   2040
            Width           =   255
         End
         Begin VB.Label Label7 
            Height          =   255
            Left            =   4920
            TabIndex        =   35
            Top             =   960
            Width           =   2655
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            Caption         =   "L"
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
            Index           =   0
            Left            =   120
            TabIndex        =   34
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            Caption         =   "R"
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
            Index           =   0
            Left            =   480
            TabIndex        =   33
            Top             =   480
            Width           =   375
         End
      End
      Begin VB.Label Label11 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4440
         TabIndex        =   131
         Top             =   510
         Width           =   3015
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Detail level:"
         Height          =   255
         Left            =   -74640
         TabIndex        =   77
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.Label Label14 
      Caption         =   "Rom Name:"
      Height          =   255
      Index           =   1
      Left            =   3720
      TabIndex        =   137
      Top             =   2880
      Width           =   975
   End
   Begin VB.Menu menu_file 
      Caption         =   "&File"
      Begin VB.Menu menu_open_rom 
         Caption         =   "&Open ROM"
         Shortcut        =   ^O
      End
      Begin VB.Menu Menu_Patch_ROM 
         Caption         =   "&Patch ROM"
         Shortcut        =   ^P
      End
      Begin VB.Menu gap6 
         Caption         =   "-"
      End
      Begin VB.Menu menu_import 
         Caption         =   "&Import"
      End
      Begin VB.Menu menu_export 
         Caption         =   "&Export"
         Begin VB.Menu menu_export_complete 
            Caption         =   "&Complete Patch File"
            Enabled         =   0   'False
         End
         Begin VB.Menu menu_export_all 
            Caption         =   "&All Changes"
            Enabled         =   0   'False
         End
         Begin VB.Menu menu_export_credits 
            Caption         =   "&Credits"
         End
         Begin VB.Menu gap7 
            Caption         =   "-"
         End
         Begin VB.Menu menu_export_level 
            Caption         =   "Current &Level (Everything)"
         End
         Begin VB.Menu menu_export_blockmap 
            Caption         =   "Current Block &Map"
         End
         Begin VB.Menu menu_export_blockdata 
            Caption         =   "Current &Block data"
         End
         Begin VB.Menu menu_sprites 
            Caption         =   "Current &Sprites"
         End
         Begin VB.Menu menu_export_tilemap 
            Caption         =   "Current &Tile-map"
            Enabled         =   0   'False
         End
         Begin VB.Menu menu_export_stops 
            Caption         =   "Current B&orders"
         End
         Begin VB.Menu gap8 
            Caption         =   "-"
         End
         Begin VB.Menu menu_export_log 
            Caption         =   "Status &Log"
         End
      End
      Begin VB.Menu gap5 
         Caption         =   "-"
      End
      Begin VB.Menu menu_save_images 
         Caption         =   "Save &Images"
         Begin VB.Menu menu_grab 
            Caption         =   "Current &Level"
         End
         Begin VB.Menu menu_grab_tilemap 
            Caption         =   "Current &Tile-map"
         End
         Begin VB.Menu menu_grab_blocks 
            Caption         =   "Current &Block-set"
         End
         Begin VB.Menu menu_grab_credits 
            Caption         =   "&Credits"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu gap3 
         Caption         =   "-"
      End
      Begin VB.Menu menu_exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu menu_edit 
      Caption         =   "&Edit"
      Begin VB.Menu menu_revert 
         Caption         =   "&Revert"
         Begin VB.Menu Menu_Revert_Map 
            Caption         =   "&Level"
         End
         Begin VB.Menu Menu_Revert_BlockData 
            Caption         =   "&Block-Data"
         End
         Begin VB.Menu Menu_Revert_Baddies 
            Caption         =   "&Sprites"
         End
         Begin VB.Menu menu_revert_credits 
            Caption         =   "&Credits"
         End
         Begin VB.Menu menu_revert_tilemap 
            Caption         =   "&TileMap"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu menu_pref 
         Caption         =   "&Preferences"
      End
   End
   Begin VB.Menu menu_view 
      Caption         =   "&View"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu menu_help 
      Caption         =   "&Help"
      Begin VB.Menu menu_howto 
         Caption         =   "Howto &Guide"
         Shortcut        =   {F1}
      End
      Begin VB.Menu menu_examples 
         Caption         =   "E&xamples"
      End
      Begin VB.Menu menu_about 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lastX As Integer, LastY As Integer, Selected_Sprite As Integer

Public Selected_Tool As Integer, Selected_Block As Integer, Map_Zoom As Integer, Display_Locked As Boolean, sprites_Loaded As Boolean, CL As Integer, Selected_Sprite_Type As Integer
Private Pattern(10, 8) As Integer, Pattern_Width As Integer, Pattern_Height As Integer, Pattern_X As Integer, Pattern_Y As Integer

Private LayerDC(4) As Long, PathInfoDC As Long
Private RedPen, BluePen

Private Sub Blocks1_MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Blocks X, Y, 2, 1
End Sub

Private Sub Blocks1_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Blocks X, Y, 2, 2
End Sub

Private Sub Blocks1_MouseUp(button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Blocks X, Y, 2, 3
End Sub
Private Sub Blocks2_MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Blocks X, Y, 1, 1
End Sub

Private Sub Blocks2_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Blocks X, Y, 1, 2
End Sub

Private Sub Blocks2_MouseUp(button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Blocks X, Y, 1, 3
End Sub
Private Sub Mouse_Blocks(X As Single, Y As Single, Optional mode As Integer, Optional Down As Integer)
block_x = Int(X / 16)
block_y = Int(Y / 16)

If Down = 1 Then
    Shape1.Tag = "m"
ElseIf Down = 3 Then
        Shape1.Tag = vbNullString
        'If button = 1 Then Selected_Block = block_y * 16 + block_x
        Shape1.Visible = True
        Shape1.Visible = True
        Load_Blocks CLng(Store(CL).Block_Bank), 114688
End If
If X < 0 Or Y < 0 Or X > Blocks2.ScaleWidth - 1 Or Y > Blocks2.ScaleHeight - 1 Then Exit Sub

If Shape1.Tag = "m" Then
    If mode = 1 Then
        Shape1.Left = block_x * 16
        Shape1.Top = block_y * 16
        Selected_Block = block_y * 16 + block_x
    ElseIf mode = 2 Then
        Tile_X = Int(X / 8)
        Tile_Y = Int(Y / 8)
        block_x = Int(X / 16)
        block_y = Int(Y / 16)
        newtile = (Shape2.Top / 8) * 16 + (Shape2.Left / 8) + 128
        If newtile > 255 Then newtile = newtile - 256
        'Shape1.Visible = False
        BitBlt Blocks1.hdc, Tile_X * 8, Tile_Y * 8, 8, 8, Tiles1.hdc, Shape2.Left, Shape2.Top, SRCCOPY
        Mid(Store(CL).Block_data, (block_y * 16 + block_x) * 4 + (Tile_Y Mod 2) * 2 + (Tile_X Mod 2) + 1, 1) = Chr$(newtile)

        Blocks1.Refresh
    End If
End If

End Sub

Private Sub Check1_Click(Index As Integer)
If Index = 0 Then
    'Check3(0) = Check1(Index)
    Check8(1) = Check1(Index)
    If Check1(Index) = True Then
        Load_Level CL
    Else
        If Main.Visible = True Then Display_Layers
    End If
Else
    If Main.Visible = True Then Load_Level CL, 2
End If


End Sub

Private Sub Check2_Click(Index As Integer)
If Index = 2 Or Index = 3 Then
    If Check2(2).value = 0 Then
        If Index = 2 Then Check2(3).value = 0
        Check1(3).Enabled = False: Check1(3).value = 0
        Check1(1).Enabled = False: Check1(1).value = 0
        Check2(3).Enabled = False: Check2(3).value = 0
        Check3(3).Enabled = False: Check3(3).value = 0
        Check3(0).Enabled = False: Check3(0).value = 0
    Else
        Check1(3).Enabled = True
        Check1(1).Enabled = True
        Check2(3).Enabled = True
        Check3(3).Enabled = True
        Check3(0).Enabled = True
    End If
    
    Check3(Index).value = Check2(Index).value

    If Main.Visible = True Then Load_Level CL, 2
End If
End Sub

Private Sub Check3_Click(Index As Integer)
If Index = 0 Then
    Check1(1) = Check3(0)
    Exit Sub
End If
Check2(Index).value = Check3(Index).value
End Sub

Private Sub Check4_Click()
Display_Credits
End Sub

Private Sub Check8_Click(Index As Integer)
Dim Checked As Boolean
Checked = Check8(Index).value
Select Case Index
Case 1
    Check1(0) = Check8(1)
Case 2
    If Checked Then
        Label7.Visible = True
        MiniMap1.height = 715
        MiniMap1.Visible = True
    Else
        Label7.Visible = False
        MiniMap1.height = 0
        MiniMap1.Visible = False
    End If
    Form_Resize
Case 3
    Display_Layers
End Select
End Sub

Private Sub Check9_Click(Index As Integer)
Display_Tile_Map Combo3.ListIndex + 1
End Sub

Private Sub Combo1_Click()
On Error GoTo oops
Combo2.Enabled = True
Combo2.Clear
Combo2.Text = "Level"
total_levels = From_ini("zone" & Combo1.ListIndex + 1, "levels", -1, True)
z = 1
Do
    levelnum = From_ini("zone" & Combo1.ListIndex + 1, z & "num", -1, True)
    If total_levels = -1 And levelnum = -1 Then GoTo done
    levelname = From_ini("maps", From_ini("zone" & Combo1.ListIndex + 1, z & "num"))
    If levelnum = -1 Then
        levelname = "Submap " & z
    Else
        If levelname = vbNullString Then levelname = "Map"
        levelname = levelname & " - " & levelnum
    End If
    
    Combo2.AddItem levelname, z - 1
z = z + 1
Loop Until total_levels > 0 And z > total_levels
done:
Combo2.ListIndex = 0
Exit Sub

oops:
Error_Handler "Combo1", Err.Description, Err.Number
End Sub

Private Sub combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo2_Click()
Dim levelnum As Integer
Combo2.Locked = True
levelnum = From_ini("zone" & Combo1.ListIndex + 1, Combo2.ListIndex + 1 & "num", -1, True)
Load_Level levelnum, 1
Auto_Center
Combo2.Locked = False
End Sub

Private Sub Combo3_Click()
set_palette palette2(0).BackColor, palette2(1).BackColor, palette2(2).BackColor, palette2(3).BackColor
Display_Tile_Map Combo3.ListIndex + 1

End Sub

Private Sub Combo4_Click()
Refresh_Status
End Sub

Private Sub Command1_Click()
Credits = Mid(ROM_Data, 431823, 1797)
Uncompressed_Credits = vbNullString
credits_Text1 = vbNullString
Display_Credits
End Sub

Private Sub Command2_Click()
Uncompressed_Credits = Parse_Credit_Text(credits_Text1)
Display_Credits
'DoEvents
'Compress_Credits
End Sub

Private Sub Compress_Credits()
temp = vbNullString
Do
    linetemp = Mid$(Credits, pos + 1, 20)
    If linetemp = String(20, " ") Then
        temp = temp & Chr(255)
    Else
        temp = temp & linetemp
        If Len(linetemp) <> 20 Then GoTo done
    End If
    pos = pos + 20
Loop
done:
Credits = temp
End Sub
Private Sub Command3_Click()
Status_data = vbNullString
status_text = vbNullString
Status "Log: Cleared", 1
'Refresh_Status False, True
End Sub

Private Sub Command4_Click()
Uncompressed_Credits = Parse_Credit_Text(credits_Text1)
Compress_Credits
DoEvents
Display_Credits
End Sub

Private Sub Command5_Click()
Display_Tile_Map Combo3.ListIndex + 1

If Paths(Text4).Loading = False Then
    Paths(Text4).X = 5
    Paths(Text4).Y = 5
End If
Load_Paths Val(Text4)
Text4 = Text4 + 1
tilemap2.Refresh
DoEvents
End Sub

Private Sub credits_Text1_KeyPress(KeyAscii As Integer)
menu_revert_credits.Enabled = True
Command2.Enabled = True
Uncompressed_Credits = vbNullString

If KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii = 32 Or KeyAscii >= 48 And KeyAscii <= 57 Then
    'allowed characters
ElseIf KeyAscii = 3 Or KeyAscii = 22 Then
    'copy and paste
ElseIf KeyAscii >= 97 And KeyAscii <= 122 Then
    KeyAscii = KeyAscii - 32
ElseIf KeyAscii = 8 Or KeyAscii = 13 Then
    'allowed characters
Else
    KeyAscii = 32
    Beep
End If
End Sub

Private Sub credits_tiles1_MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
credits_tiles1_mouse 0, button, X, Y
End Sub

Private Sub credits_tiles1_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
credits_tiles1_mouse 1, button, X, Y
End Sub

Private Sub credits_tiles1_MouseUp(button As Integer, Shift As Integer, X As Single, Y As Single)
credits_tiles1_mouse 2, button, X, Y
End Sub

Private Sub credits_tiles1_mouse(action As Integer, button As Integer, X As Single, Y As Single)
Select Case action
Case 0
    Shape6.Tag = "m"
Case 1

Case 2
    Shape6.Tag = vbNullString
End Select

If Shape6.Tag = "m" Then
    sx = Int(X / 8)
    sy = Int(Y / 8)
    Shape6.Left = sx * 8
    Shape6.Top = sy * 8
End If
End Sub

Private Sub credits_VScroll1_Change()
credits1.Top = -credits_VScroll1
End Sub

Private Sub credits_VScroll1_Scroll()
credits1.Top = -credits_VScroll1
End Sub

Private Sub credits1_Resize()
credits_VScroll1.Max = credits1.height
End Sub




Private Sub Credits1_mouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Credits1 0, button, X, Y
End Sub

Private Sub Credits1_mouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Credits1 1, button, X, Y
End Sub

Private Sub Credits1_mouseUp(button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Credits1 2, button, X, Y
End Sub

Private Sub Mouse_Credits1(action As Integer, button As Integer, X As Single, Y As Single)
Command2.Enabled = True
menu_revert_credits.Enabled = True

Select Case action
Case 0
    credits1.Tag = "m"
Case 2
    credits1.Tag = vbNullString
End Select

If X < 0 Or Y < 0 Or X > credits1.ScaleWidth - 1 Or Y > credits1.ScaleHeight - 1 Then Exit Sub

If credits1.Tag = "m" Then
    sx = Int(X / 8)
    sy = Int(Y / 8)
    block_x = Int(Shape6.Left / 8)
    block_y = Int(Shape6.Top / 8)
    
    BitBlt credits1.hdc, sx * 8, sy * 8, 8, 8, credits_tiles1.hdc, Shape6.Left, Shape6.Top, SRCCOPY
    Mid(Uncompressed_Credits, sx + (sy * 20) + 1, 1) = Chr((block_y * 16) + block_x)
        
    credits1.Refresh
End If
End Sub

Private Sub Load_credits_tiles()
Dim X As Integer, Y As Integer
For X = 0 To 15
    For Y = 0 To 15
        General_Load_Tile credits_tiles1, 426839, Y * 16 + X, X, Y
    Next Y
Next X
credits_tiles1.Line (0, 16)-(8, 24), vbWhite, BF
credits_tiles1.Refresh
End Sub

Private Sub Decompress_Credits()
If Uncompressed_Credits <> vbNullString Then Exit Sub
Do
    For X = 0 To 19
        tileasc = Asc(Mid(Credits, byte_pos + 1, 1))
        If tileasc = 255 Then
            Uncompressed_Credits = Uncompressed_Credits & String(20, " ")
        Else
            Uncompressed_Credits = Uncompressed_Credits & Chr(tileasc)
        End If
        byte_pos = byte_pos + 1
        If byte_pos >= Len(Credits) Then Exit For
    Next X
Loop Until byte_pos >= Len(Credits)
End Sub

Private Sub Display_Credits()
Dim Credits_Text_Temp As String
If Len(Uncompressed_Credits) = 0 Then
    If Len(credits_Text1) > 0 Then
        Uncompressed_Credits = Parse_Credit_Text(credits_Text1)
    Else
        Decompress_Credits
    End If
End If
credits1.height = 10
credits1.Cls
credits_Text1 = vbNullString
Do
    credits1.height = (Y + 1) * 8
    For X = 0 To 19
        'If byte_pos + 1 > Len(Uncompressed_Credits) Then GoTo end_line
        tileasc = Asc(Mid(Uncompressed_Credits, byte_pos + 1, 1))
        Credits_Text_Temp = Credits_Text_Temp & Chr(tileasc)
        sy = Int(tileasc / 16)
        sx = tileasc - sy * 16
        
        BitBlt credits1.hdc, X * 8, Y * 8, 8, 8, credits_tiles1.hdc, sx * 8, sy * 8, SRCCOPY

        byte_pos = byte_pos + 1
        If byte_pos >= Len(Uncompressed_Credits) Then Exit For
    Next X
'end_line:
    Y = Y + 1
    Credits_Text_Temp = Credits_Text_Temp & vbCrLf
    If Y Mod 4 = 0 Then
        Credits_Text_Temp = Credits_Text_Temp & "<FADE>" & vbCrLf
        If Check4.value = vbChecked Then credits1.Line (0, Y * 8 - 1)-(20 * 8, Y * 8 - 1), RGB(255, 0, 0)
    End If
    'credits1.height = credits1.height + 8
    line_num = line_num + 1
Loop Until byte_pos >= Len(Uncompressed_Credits)
credits_Text1 = Credits_Text_Temp
credits1.Refresh
DoEvents
End Sub

Private Sub Form_Load()
On Error GoTo oops
If App.PrevInstance = True Then
    SmellyMoo = True
    Unload Me
End If

Status "Log: Started", 2
MiniMap1.BackColor = SSTab1.BackColor
Map_Zoom = 16
change_tool 0
Me.Caption = Main.Caption & " (" & App.Major & "." & App.Minor & ")"
SSTab1.Tab = 0: SSTab2.Tab = 0
Combo4.ListIndex = 2

CreateBuffers

If Len(Command$) > 0 Then
    Dim temp As String
    
    temp = Replace(LCase$(Command$), Chr(34), vbNullString)
    If Right(temp, 3) = ".gb" Then
        Load_ROM temp
    ElseIf Right(temp, 4) = ".mpf" Then
        Load_ROM From_ini("config", "rom")
        Import_File temp
    End If
Else
    Load_ROM From_ini("config", "rom")
End If

'###CREDITS:
Credits = Mid(ROM_Data, 431823, 1797) '1797

Exit Sub

oops:
Error_Handler "Form_Load", Err.Description, Err.Number
End Sub

Private Sub CreateBuffers()
Dim OldBmp As Long, LayerBMP(4) As Long, PathInfoBMP As Long
For L = 0 To UBound(LayerDC)
    LayerDC(L) = CreateCompatibleDC(TileMap1.hdc)
    LayerBMP(L) = CreateCompatibleBitmap(TileMap1.hdc, 256 * 16, 48 * 16)
    OldBmp = SelectObject(LayerDC(L), LayerBMP(L))
    DeleteObject OldBmp
    DeleteObject LayerBMP(L)
Next L

PathInfoDC = CreateCompatibleDC(tilemap2.hdc)
PathInfoBMP = CreateCompatibleBitmap(tilemap2.hdc, 128 * 8, 128 * 8)
OldBmp = SelectObject(PathInfoDC, PathInfoBMP)
DeleteObject OldBmp
DeleteObject PathInfoBMP
    
'BitBlt LayerDC(2), 0, 0, 256 * 16, 48 * 16, 0, 0, 0, vbWhiteness
'BitBlt LayerDC(3), 0, 0, 256 * 16, 48 * 16, 0, 0, 0, vbWhiteness
RedPen = CreatePen(PS_SOLID, 1, vbRed)
DeleteObject SelectObject(LayerDC(3), RedPen)
End Sub
Private Sub deleteBuffers()
For L = 0 To UBound(LayerDC)
    DeleteDC (LayerDC(L))
Next L
DeleteDC PathInfoDC
DeleteObject (RedPen)
End Sub
Private Sub DrawGDILine(hdc As Long, x1, y1, x2, y2)
MoveToEx hdc, x1, y1, 0
LineTo hdc, x2, y2
End Sub

Private Sub ColourBuffer(hdc As Long, width As Integer, height As Integer, colour As Long)
Dim R As RECT, Pen As Long, OldPen As Long
SetRect R, 0, 0, width, height

Pen = CreatePen(PS_SOLID, 1, colour)
OldPen = SelectObject(hdc, Pen)
FillRect hdc, R, Pen
DeleteObject SelectObject(hdc, OldPen)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If SmellyMoo Then Exit Sub
response = MsgBox("Are you sure?" & vbCrLf & vbCrLf & "unsaved changes will be lost.", vbExclamation + vbOKCancel, "Exit?")
If response <> 1 Then Cancel = True
End Sub

Private Sub Form_Resize()
On Error GoTo oops
If Me.WindowState = 1 Then Exit Sub
If Me.width < 7000 Then Me.width = 7000
If Me.height < 6700 Then Me.height = 6700

SSTab2.width = Me.width - 65
SSTab2.height = Me.height - SSTab2.Top - 810

SSTab1.width = SSTab2.width
SSTab1.height = SSTab2.height - SSTab1.Top + 10

SSTab3.width = SSTab2.width
SSTab3.height = SSTab2.height - SSTab3.Top + 10

Select Case SSTab2.Tab
Case 0 'level
    Select Case SSTab1.Tab
    Case 0 'edit
        Picture2.Top = MiniMap1.Top + MiniMap1.height + 240
        Picture2.width = SSTab1.width - Picture2.Left - 420
        VScroll1.Top = Picture2.Top
        VScroll1.Left = Picture2.Left + Picture2.width
        HScroll1.width = Picture2.width
        If Blocks2.Visible = True Then
            Blocks2.Top = SSTab1.height - Blocks2.height - 240
            Pattern1.Top = Blocks2.Top
            Picture2.height = SSTab1.height - Picture2.Top - (SSTab1.height - Blocks2.Top) - 420
        ElseIf Picture3.Visible = True Then
            Picture3.Top = SSTab1.height - Picture3.height - 240
            VScroll3.Top = Picture3.Top
            Picture2.height = SSTab1.height - Picture2.Top - (SSTab1.height - Picture3.Top) - 420
        Else
            Picture2.height = SSTab1.height - Picture2.Top - 420
        End If
        HScroll1.Top = Picture2.Top + Picture2.height
        VScroll1.height = Picture2.height
    Case 2 ' sprites
            Picture1.width = SSTab1.width - Picture1.Left - 660
            If Picture1.ScaleWidth >= sprites2.width Then Picture1.width = sprites2.width * Screen.TwipsPerPixelX
            Picture1.height = SSTab1.height - Picture1.Top - 420
            VScroll2.Left = Picture1.Left + Picture1.width
            VScroll2.height = Picture1.height
            HScroll2.width = Picture1.width
            HScroll2.Top = Picture1.Top + Picture1.height
            'Frame4.Left = Picture1.Left + Picture1.Width + 420
    End Select
Case 1 'global
    Select Case SSTab3.Tab
    Case 0 'tilemap
        If SSTab3.Tab = 0 Then
            tilemap2_Resize
            Picture4.width = SSTab3.width - Picture4.Left - 660 - Frame6.width
            Picture4.height = SSTab3.height - Picture4.Top - 420
            tilemap_VScroll.Left = Picture4.Left + Picture4.width
            tilemap_VScroll.height = Picture4.height
            tilemap_HScroll.width = Picture4.width
            tilemap_HScroll.Top = Picture4.Top + Picture4.height
            Frame6.Left = Picture4.Left + Picture4.width + 420
        End If
    Case 1 'text
        credits_Text1.height = SSTab3.height - credits_Text1.Top - 660
        Command1.Top = credits_Text1.Top + credits_Text1.height + 240
        Command2.Top = Command1.Top
        Command4.Top = Command1.Top
    Case 2 ' graphics
        credits_Picture1.height = SSTab3.height - credits_Picture1.Top - 420
        credits_VScroll1.height = credits_Picture1.height
    End Select
Case 2 'config
    'nop
Case 3 'status
    Command3.Top = SSTab2.height - Command3.height - 240
    status_text.width = SSTab2.width - 420
    status_text.height = Command3.Top - status_text.Top - 240
End Select

'DoEvents
Exit Sub

oops:
Error_Handler "Form_Resize", Err.Description, Err.Number
End Sub

Private Sub Form_Unload(Cancel As Integer)
deleteBuffers
Close 1
End
End Sub

Private Sub HScroll1_Change()
TileMap1.Left = -HScroll1.value - 1
'Shape4.Left = HScroll1.Value / Map_Zoom
Display_ScreenPos
End Sub

Private Sub HScroll2_Change()
sprites2.Left = -HScroll2.value
End Sub

Private Sub HScroll2_Scroll()
HScroll2_Change
End Sub

Private Sub menu_about_Click()
About.Visible = True
End Sub

Private Sub menu_credits_Click()
Credit.Show
End Sub

Private Sub menu_examples_Click()
On Error Resume Next
ShellExecute Me.hWnd, "open", App.Path & "\help\examples", "", "", 1
End Sub

Private Sub menu_exit_Click()
End
End Sub

Private Sub menu_findsub_Click(Index As Integer)
Find.Visible = True
Find.SSTab1.Tab = Index
End Sub

Private Sub menu_export_blockdata_Click()
Export_File Store(CL).Block_data, "Block", , CL + 1
End Sub

Private Sub menu_export_blockmap_Click()
Export_File Store(CL).Unzipped_Modified, "Map", , CL + 1
End Sub

Private Sub menu_export_credits_Click()
If Len(Uncompressed_Credits) = 0 Then
    If Len(credits_Text1) > 0 Then
        Uncompressed_Credits = Parse_Credit_Text(credits_Text1)
    Else
        Decompress_Credits
    End If
End If
Export_File Uncompressed_Credits, "Credits"
End Sub

Public Sub Export_File(Data As String, File_Type As String, Optional Default_name As String, Optional Level_num As Integer)
If Len(Data) = 0 Then Exit Sub

If Len(Default_name) = 0 Then
    If Level_num > 0 Then
        Default_name = File_Type & " " & Level_num
    Else
        Default_name = File_Type
    End If
End If

FileName = Display_export(Default_name)
If FileName = vbnullstr Then Exit Sub

Close 2
Kill FileName
Open FileName For Binary As 2

Put 2, , Export_Header(File_Type, Len(Data), Level_num) & Data
Close 2
End Sub
Public Sub Export_Files(Sections As String, Optional Default_name As String)
If Len(Sections) = 0 Then Exit Sub

If Len(Default_name) = 0 Then Default_name = "Export"

FileName = Display_export(Default_name)
If FileName = vbnullstr Then Exit Sub

Close 2
Open FileName For Binary As 2

Put 2, , Sections
Close 2
End Sub

Private Sub menu_export_level_Click()
Dim temp As String, Temp_Data As String
Temp_Data = Store(CL).Unzipped_Modified
temp = Export_Header("Map", Len(Temp_Data), CL + 1) & Temp_Data

Temp_Data = Store(CL).Borders
temp = temp & Export_Header("Borders", Len(Temp_Data), CL + 1) & Temp_Data

Save_Sprites
Temp_Data = Store(CL).Sprites
temp = temp & Export_Header("Sprites", Len(Temp_Data), CL + 1) & Temp_Data

Temp_Data = Store(CL).Block_data
temp = temp & Export_Header("Block", Len(Temp_Data), CL + 1) & Temp_Data

Export_Files temp, "Level"
End Sub

Public Sub menu_export_log_Click()
Export_File Status_data, "Log"
End Sub

Private Sub menu_export_stops_Click()
Export_File Store(CL).Borders, "Borders", , CL + 1
End Sub

Private Sub menu_grab_credits_Click()
On Error GoTo oops
With CommonDialog1
    .Filter = "Jpeg (*.jpg)|*.jpg|Gif (*.gif)|*.gif|All Files (*.*)|*.*|"
    .FileName = "credits.jpg"
    .ShowSave
    SavePicture credits1.Image, .FileName
End With
Exit Sub

oops:
If Err.Number = 32755 Then Exit Sub
MsgBox Err.Description
End Sub

Private Sub menu_grab_tilemap_Click()
On Error GoTo oops
With CommonDialog1
    .Filter = "Jpeg (*.jpg)|*.jpg|Gif (*.gif)|*.gif|All Files (*.*)|*.*|"
    .FileName = "tilemap.jpg"
    .ShowSave
    SavePicture tilemap2.Image, .FileName
End With
Exit Sub

oops:
If Err.Number = 32755 Then Exit Sub
MsgBox Err.Description
End Sub

Private Sub menu_howto_Click()
On Error Resume Next
ShellExecute Me.hWnd, "open", App.Path & "\help\help.htm", "", "", 1
End Sub

Public Sub menu_import_Click()
Import_File
End Sub

Public Sub Import_File(Optional File_Name As String)
Dim Header_Start As Integer, Header_End As Integer, temp() As String
Dim Export_Type As String, Export_Size As Integer, Export_Data As String, File_Data As String, Export_level As Integer
Dim AutoLevel As Boolean, Level As Integer
On Error GoTo oops:

If Len(File_Name) = 0 Then
    With CommonDialog1
        .MaxFileSize = 200
        .Filter = "Mario Patch Files (.mpf/.moo)|*.mpf;*.moo|All Files (*.*)|*.*|"
        .ShowOpen
        File_Name = .FileName
    End With
End If
Open File_Name For Binary As 2
Status "Importing: Started", 2
z = LOF(2)
File_Data = Space(LOF(2))

Get 2, , File_Data

Header_Start = 2
Do
    Header_End = InStr(Header_Start, File_Data, ">")
    elements = Split(Mid(File_Data, Header_Start, Header_End - Header_Start), " ")
    For Each element In elements
        temp() = Split(element, "=")
        Select Case LCase(temp(0))
        Case "type"
            Export_Type = temp(1)
        Case "size"
            Export_Size = CInt(temp(1))
        Case "level"
            If AutoLevel = False Then
                ret = MsgBox("Each part of this import has an assosiated level." & vbCrLf & "Would you like to import them there? (Y) or this level? (N)", vbYesNo, "Which level?")
                If ret = 6 Then AutoLevel = True
            End If
            Export_level = CInt(temp(1))
        'case "redirect"
        'case "protected"
        'case "next"
        End Select
    Next element
    
    Export_Data = Mid(File_Data, Header_End + 1, Export_Size)
    
    Status ", " & Export_Type, 2, True
    
    If AutoLevel Then Level = Export_level - 1 Else Level = CL
    
    Select Case LCase(Export_Type)
    Case "map"
        Store(Level).Unzipped_Modified = Export_Data
    Case "sprites"
        Store(Level).Sprites = Export_Data
        sprites_Loaded = False
    Case "credits"
        Uncompressed_Credits = Export_Data
    Case "blocks"
        Store(Level).Block_data = Export_Data
        Load_Blocks CLng(Store(Level).Block_Bank), 114688
    Case "log"
        Status_data = Export_Data
        Refresh_Status
    Case "borders"
        Store(Level).Borders = Export_Data
    Case vbNullString
        Status "Invalid Import File", 4
    End Select
Header_Start = InStr(Header_End + Export_Size, File_Data, "<") + 1
Loop Until Header_End + Export_Size >= LOF(2) Or Header_Start = 1
Load_Level CL, 1
Close 2
Exit Sub

oops:
If Err.Number = 32755 Then Exit Sub
Error_Handler "Display_export", Err.Description, Err.Number
End Sub

Private Function Display_export(export_name As String) As String
On Error GoTo oops:
With CommonDialog1
    .MaxFileSize = 200
    .Filter = "Mario Patch Files (*.moo)|*.moo|All Files (*.*)|*.*|"
    .FileName = export_name & ".moo"
    .ShowSave
    Display_export = .FileName
End With
Exit Function

oops:
If Err.Number = 32755 Then Exit Function
Error_Handler "Display_export", Err.Description, Err.Number
End Function

Private Sub menu_grab_Click()
On Error GoTo oops
With CommonDialog1
    .Filter = "Jpeg (*.jpg)|*.jpg|Gif (*.gif)|*.gif|All Files (*.*)|*.*|"
    .FileName = "map " & CL & ".jpg"
    .ShowSave
    SavePicture TileMap1.Image, .FileName
End With
Exit Sub

oops:
If Err.Number = 32755 Then Exit Sub
MsgBox Err.Description
End Sub

Private Sub menu_grap_tilemap_Click()
On Error GoTo oops
With Main.CommonDialog1
    .Filter = "Jpeg (*.jpg)|*.jpg|Gif (*.gif)|*.gif|All Files (*.*)|*.*|"
    .FileName = "Tilemap.jpg"
    .ShowSave
    SavePicture tilemap2.Image, .FileName
End With
Exit Sub

oops:
If Err.Number = 32755 Then Exit Sub
MsgBox Err.Description
End Sub

Private Sub menu_open_rom_Click()
On Error GoTo oops:
With CommonDialog1
    .MaxFileSize = 200
    .Filter = "GameBoy Rom's (*.gb)|*.gb|All Files (*.*)|*.*|"
    .ShowOpen
    Load_ROM .FileName
    WritePrivateProfileString "config", "rom", .FileName, App.Path & "\config\config.ini"
End With

If Len(Credits) = 0 Then Credits = Mid(ROM_Data, 431823, 1797)

Exit Sub

oops:
If Err.Number = 32755 Then Exit Sub
Error_Handler "Open_ROM", Err.Description, Err.Number
End Sub

Private Sub Menu_Patch_ROM_Click()
Patching.Show
End Sub

Public Sub Patch_Rom()
Dim block_location As Long, Bank_Temp As String, Level_num As Integer, B As Integer, s As Integer
Dim Temp_Sprites As String
On Error GoTo oops

Patching.Show
Status "Patching: Started", 6, , True
Status "Loading: all unloaded levels", 0, , True
Load_All_Levels

Save_Sprites

For B = 0 To UBound(Bank_Store())
    Bank_Temp = vbNullString
    Status "Compiling: Bank " & Bank_Store(B).Bank & " Maps: (levels " & Bank_Store(B).Level_list & ")", 2, , True
    For s = 0 To Bank_Store(B).Last_Sub
        total_levels = total_levels + 1
        Level_num = Find_Map(CInt(Bank_Store(B).Bank), s)
        Compress_map Level_num
        Bank_Temp = Bank_Temp & Store(Level_num).Zipped
    Next s
    Status "Compiled: Bank " & Bank_Store(B).Bank & " Maps. size: " & Len(Bank_Temp), 2, , True
    block_location = Skip_To(CInt(Bank_Store(B).Bank), 0)
    If Len(Bank_Temp) > Val(From_ini("Max", "Bank" & Bank_Store(B).Bank)) Then
        Status "Compiling: Does not fit. modded bank size= " & Len(Bank_Temp) & " (Max=" & Val(From_ini("Max", "Bank" & Bank_Store(B).Bank)) & ")", 3, , True
        Status "Block-map to large. Increase of overall horizontal detail. will probably generate corrupted ROM.", 3, , True
    Else
        Status "Compiling: Fits (Max=" & Val(From_ini("Max", "Bank" & Bank_Store(B).Bank)) & ")", 1, , True
    End If
    '", Unsafe max size: 16384"
    'Status "Info: Next map bank (Location = " & block_location & ")", 1, , True
    Mid(ROM_Data, block_location) = Bank_Temp
Next B

Status "Patching: Rom Name", 1, , True
new_name = "Mario 2 Mod"
Mid(ROM_Data, 309) = new_name & String(19 - Len(new_name), vbNullChar)

Status "Patching: Rom Version", 1, , True
Mid(ROM_Data, 333) = Chr$(Val(Text3))

For Level_num = 0 To total_levels - 1
    If Len(Store(Level_num).Block_data) > 0 Then
        Status "Patching: Block-Data (level " & Level_num & ")", 1, , True
        Mid(ROM_Data, 114688 + 256 * Store(Level_num).Block_Bank + 1, 512) = Store(Level_num).Block_data
    End If
    
    Status "Compiling: Sprites (level " & Level_num & ")", 1, , True
    Temp_Sprites = Temp_Sprites & Store(Level_num).Sprites & Chr(255)
Next Level_num

Status "Patching: Sprites", 1, , True
Mid(ROM_Data, 57463 + 1) = Temp_Sprites

Status "Verifing: " & Len(Temp_Sprites) \ 3 & " sprites in total", 1, , True
If Len(Temp_Sprites) \ 3 > 959 Then
    Status "Patching: Exceeded global maximum of sprites/items, ROM may be corrupted.", 3, , True
End If

Status "Patching: Borders", 1, , True
Save_ALl_Borders

Status "Patching: Credits", 1, , True
Mid(ROM_Data, 431823, 1797) = Left$(Credits, 1797)


Open App.Path & "\" & ROM_Name & " Patched.gb" For Binary As 3
Put 3, 1, ROM_Data
Close 3
Status "Completed: Created '..." & ROM_Name & " Patched.gb'", 6, , True
Exit Sub

oops:
Error_Handler "Patch_ROM", Err.Description, Err.Number
End Sub
Private Function Find_Map(Bank As Integer, Sub_Bank As Integer) As Integer
For L = 0 To UBound(Store())
    If Store(L).Map_Bank = Bank And Store(L).Map_Sub_Bank = Sub_Bank Then GoTo found
Next L
Find_Map = -1
Exit Function

found:
Find_Map = L
End Function


Private Sub menu_pref_Click()
If SSTab2.Enabled Then SSTab2.Tab = 2
End Sub

Private Sub Menu_Revert_Baddies_Click()
Store(CL).Sprites = vbNullString
sprites_Loaded = False
End Sub

Private Sub Menu_Revert_BlockData_Click()
Store(CL).Block_data = vbNullString
Load_Blocks CLng(Store(CL).Block_Bank), 114688
End Sub

Private Sub menu_revert_Click()
If CL < 0 Then
    Menu_Revert_Map.Enabled = False
ElseIf Store(CL).Unzipped <> Store(CL).Unzipped_Modified And Len(Store(CL).Unzipped) Then
    Menu_Revert_Map.Enabled = True
Else
    Menu_Revert_Map.Enabled = False
End If
End Sub

Private Sub menu_revert_credits_Click()
Credits = Mid(ROM_Data, 431823, 1797)
Uncompressed_Credits = vbNullString
credits_Text1 = vbNullString
DoEvents
Display_Credits
End Sub

Private Sub Menu_Revert_Map_Click()
Store(CL).Unzipped_Modified = Store(CL).Unzipped
Load_Level CL, 0
End Sub

Private Sub menu_sprites_Click()
Save_Sprites
Export_File Store(CL).Sprites, "Sprites", , CL + 1
End Sub

Private Sub MiniMap1_MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
MiniMap1.Tag = "m"
Center_On CInt(X), CInt(Y)
End Sub

Private Sub MiniMap1_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
If MiniMap1.Tag = "m" Then
    Center_On CInt(X), CInt(Y)
End If
End Sub

Private Sub MiniMap1_MouseUp(button As Integer, Shift As Integer, X As Single, Y As Single)
If MiniMap1.Tag = "m" Then
    Center_On CInt(X), CInt(Y)
End If
MiniMap1.Tag = vbNullString
End Sub

Private Sub palette1_Click(Index As Integer)
On Error GoTo oops
CommonDialog1.ShowColor
palette1(Index).BackColor = CommonDialog1.Color
Current_Palette(Index) = CommonDialog1.Color
Load_Level CL, 1
Exit Sub

oops:
If Err.Number <> 32755 Then Error_Handler "Pallette click", Err.Description, Err.Number
End Sub

Private Sub palette2_Click(Index As Integer)
On Error GoTo oops
CommonDialog1.ShowColor
palette2(Index).BackColor = CommonDialog1.Color
Current_Palette(Index) = CommonDialog1.Color

Display_Tile_Map Combo3.ListIndex + 1
Exit Sub

oops:
If Err.Number <> 32755 Then Error_Handler "Pallette click", Err.Description, Err.Number
End Sub

Private Sub palette3_Click(Index As Integer)
On Error GoTo oops
CommonDialog1.ShowColor
palette3(Index).BackColor = CommonDialog1.Color
Current_Palette(Index) = CommonDialog1.Color
DoEvents

Load_credits_tiles
DoEvents
Display_Credits
Exit Sub

oops:
If Err.Number <> 32755 Then Error_Handler "Pallette click", Err.Description, Err.Number
End Sub

Private Sub Pattern1_MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
Dim mX As Integer, mY As Integer
mX = Int(X / 16)
mY = Int(Y / 16)
Pattern1.Tag = "m"
If button = 1 Then
    set_pattern_block Selected_Block, mX, mY
ElseIf button = 2 Then
    'Pattern(mX, mY) = 0
    set_pattern_block -1, mX, mY
End If
scale_pattern
End Sub

Private Sub Pattern1_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
If Pattern1.Tag = "m" Then
    Dim mX As Integer, mY As Integer
    mX = Int(X / 16)
    mY = Int(Y / 16)
    If mX < 0 Or mY < 0 Then Exit Sub
    If button = 1 Then
        set_pattern_block Selected_Block, mX, mY
    ElseIf button = 2 Then
        'If mX <= 10 And mY <= 8 Then
            set_pattern_block -1, mX, mY
            'Pattern(mX, mY) = 0
        'End If
    End If
End If
scale_pattern
End Sub

Private Sub Pattern1_MouseUp(button As Integer, Shift As Integer, X As Single, Y As Single)
Dim mX As Integer, mY As Integer
mX = Int(X / 16)
mY = Int(Y / 16)
Pattern1.Tag = vbNullString
If button = 1 Then
    set_pattern_block Selected_Block, mX, mY
ElseIf button = 2 Then
    set_pattern_block -1, mX, mY
    'Pattern(mX, mY) = 0
End If
scale_pattern
End Sub

Private Sub set_pattern_block(block_type As Integer, mX As Integer, mY As Integer)
block_y = Int(block_type / 16)
block_x = block_type - block_y * 16
BitBlt Pattern1.hdc, mX * 16, mY * 16, 16, 16, Blocks1.hdc, block_x * 16, block_y * 16, SRCCOPY
If mX <= 10 And mX >= 0 And mY <= 8 And mY >= 0 Then
    Pattern(mX, mY) = block_type + 1
End If
Pattern1.Refresh
End Sub

Private Sub pattern_tool(mX As Integer, mY As Integer)
Dim block_type As Integer
If mX < 0 Or mY < 0 Or (Pattern_Width = 0 And Pattern_Height = 0) Then Exit Sub
px = (mX - Pattern_X) Mod Pattern_Width: py = (mY - Pattern_Y) Mod Pattern_Height
If px < 0 Then px = px + Pattern_Width
If py < 0 Then py = py + Pattern_Height
block_type = Pattern(px, py) - 1
If block_type = -1 Then Exit Sub
block_y = Int(block_type / 16)
block_x = block_type - block_y * 16
change_map mX, mY, block_type, True
End Sub
Private Sub scale_pattern()
Pattern_Width = 0: Pattern_Height = 0
For TX = 0 To 9
    For TY = 0 To 7
        If Pattern(TX, TY) > 0 Then
            If TX + 1 > Pattern_Width Then Pattern_Width = TX + 1
            If TY + 1 > Pattern_Height Then Pattern_Height = TY + 1
        Else
            Pattern1.Line (TX * 16, TY * 16)-((TX + 1) * 16 - 1, (TY + 1) * 16 - 1), Pattern1.ForeColor, BF
        End If
    Next TY
Next TX
    
Pattern1.Line (Pattern_Width * 16, 0 * 16)-(10 * 16, 8 * 16), Pattern1.BackColor, BF
Pattern1.Line (0 * 16, Pattern_Height * 16)-(10 * 16, 8 * 16), Pattern1.BackColor, BF
Pattern1.Refresh
End Sub

Private Sub Picture1_Resize()
xover = sprites2.width - Picture1.ScaleWidth - 2
yover = sprites2.height - Picture1.ScaleHeight - 2
If yover > 0 Then
    VScroll2.Visible = True
    VScroll2.Max = yover
Else
    VScroll2.Visible = False
End If
If xover > 0 Then
    HScroll2.Visible = True
    HScroll2.Max = xover
Else
    HScroll2.Visible = False
End If

End Sub

Private Sub Picture2_Resize()
ScrollBars
End Sub

Private Sub reg_Timer()

End Sub

Private Sub sprites3_MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
sprite_x = Int(X / 32)
sprite_y = Int(Y / 32)

Selected_Sprite_Type = sprite_y * 8 + sprite_x

change_selected_sprite
If Selected_Sprite = -1 Then Exit Sub
sprite_Store(Selected_Sprite).typeA = sprite_y
sprite_Store(Selected_Sprite).TypeB = sprite_x
Load_Level CL, 0
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Select Case SSTab1.Tab
Case 0
    If Len(ROM_Name) > 0 Then
        If Main.Visible = True Then Load_Level CL, 0
    End If
End Select
Form_Resize
End Sub

Private Sub SSTab2_Click(PreviousTab As Integer)
If ROM_Name = vbNullString Then Exit Sub
Form_Resize
Select Case SSTab2.Tab
Case 0 'level
    set_palette palette1(0).BackColor, palette1(1).BackColor, palette1(2).BackColor, palette1(3).BackColor
    If Main.Visible = True And SSTab1.Tab = 0 Then Load_Level CL, 0
    SSTab1.Tab = 0
Case 1 'global
    SSTab3.Tab = 0
    
    Combo3.Enabled = True
    Combo3.Clear
    Combo3.Text = "Select Map"
    
    total_levels = From_ini("overworld", "total", -1, True)
    For z = 1 To total_levels
        map_name = From_ini("overworld", z & "name", "error")
        Combo3.AddItem map_name, z - 1
    Next z
    
    If Combo3.ListCount > 0 Then Combo3.ListIndex = 0
    
    'DoEvents
End Select
End Sub

Private Sub SSTab3_Click(PreviousTab As Integer)
Form_Resize
Select Case SSTab3.Tab
Case 1 'credits text
    DoEvents
    Decompress_Credits
    Display_Credits
Case 2 'credits image
    set_palette palette3(0).BackColor, palette3(1).BackColor, palette3(2).BackColor, palette3(3).BackColor
    'DoEvents
    Load_credits_tiles
    DoEvents
    Display_Credits
    Command2.Enabled = False
    menu_grab_credits.Enabled = True
End Select
End Sub

Private Sub tilemap_HScroll_Change()
tilemap2.Left = -tilemap_HScroll.value - 1
End Sub

Private Sub tilemap_HScroll_Scroll()
tilemap2.Left = -tilemap_HScroll.value - 1
End Sub

Private Sub tilemap_VScroll_Change()
tilemap2.Top = -tilemap_VScroll.value - 1
End Sub

Private Sub tilemap_VScroll_Scroll()
tilemap2.Top = -tilemap_VScroll.value - 1
End Sub

Private Sub TileMap1_MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo oops
Dim mX As Integer, mY As Integer
mX = Int(X / Map_Zoom)
mY = Int(Y / Map_Zoom)
TileMap1.Tag = "m"
If button = 2 Then
    temp_tool = Selected_Tool + 1
Else
    temp_tool = Selected_Tool
End If
Select Case temp_tool
Case 0, 9 'pen
    change_map mX, mY, Selected_Block, True
Case 1, 5, 7 'sampe
    sample_tool mX, mY
    Exit Sub
Case 2 'brush
    If Shift = 0 Then
        pattern_tool mX, mY
    Else
        Pattern_X = mX: Pattern_Y = mY
        pattern_tool mX, mY
    End If
Case 3, 4 ' select
    Shape3.Visible = True
    Shape3.Left = mX * Map_Zoom
    Shape3.Top = mY * Map_Zoom
    Shape3.width = Map_Zoom: Shape3.height = Map_Zoom
    lastX = mX
    LastY = mY
Case 8 'erase
    Selected_byte = Mid(Store(CL).Unzipped, mY * 256 + mX + 1)
    change_map mX, mY, Asc(Selected_byte), True
Case 10, 11 'sprites
    Selected_Sprite = -1
    Shape3.width = Map_Zoom
    Shape3.height = Map_Zoom
    For pos = 0 To UBound(sprite_Store())
        If sprite_Store(pos).X - 1 > mX * 2 - 2 And sprite_Store(pos).X - 1 < mX * 2 + 4 _
        And sprite_Store(pos).Y - 2 > mY * 2 - 2 And sprite_Store(pos).Y - 2 < mY * 2 + 4 Then
            Selected_Sprite = pos
            Shape3.Visible = True
            Shape3.Left = Int(X / Map_Zoom * 2) * Map_Zoom / 2
            Shape3.Top = (Int(Y / Map_Zoom * 2)) * Map_Zoom / 2
            Selected_Sprite_Type = sprite_Store(pos).typeA * 8 + sprite_Store(pos).TypeB
            change_selected_sprite
            Exit For
        End If
    Next pos
    If sprite_Store(0).X < X + 32 And sprite_Store(0).X > X - 16 And _
    sprite_Store(0).Y < Y + 32 And sprite_Store(0).Y > Y - 32 Then
        Selected_Sprite = 0
        Shape3.Visible = True
        Shape3.Left = X
        Shape3.Top = Y
    End If
Case 12, 13 'mag
    If button = 2 Then
        Map_Zoom = Map_Zoom \ 2
    ElseIf button = 1 Then
        Map_Zoom = Map_Zoom * 2
    End If
    If Map_Zoom < 1 Then
        Map_Zoom = 1
    ElseIf Map_Zoom > 32 Then
        Map_Zoom = 32
    Else
        TileMap1.Visible = False
        Display_Layers
    End If
Case 14 'borders
    Dim t, B, L, R As Boolean, Modify As Byte
    sbx = X \ 256 '(16*16)
    sby = Y \ 256
    t = (Y - sby * 256 < 64)
    B = (Y - sby * 256 > 192)
    L = (X - sbx * 256 < 64)
    R = (X - sbx * 256 > 192)
    If B Then Modify = 8
    If t Then Modify = Modify + 4
    If L Then Modify = Modify + 2
    If R Then Modify = Modify + 1
    Modify = Asc(Mid(Store(CL).Borders, sby * 16 + sbx + 1, 1)) Xor Modify
    Mid(Store(CL).Borders, sby * 16 + sbx + 1, 1) = Chr(Modify)
    Display_Borders CL
    Display_Layers
Case 15 ' warps
    sbx = X \ 256 '(16*16)
    sby = Y \ 256
    warp = Asc(Mid(Store(CL).Warps, sby * 16 + sbx + 1, 1))
    MsgBox "Warps editor needs work. warp code for this square=" & warp
End Select
Exit Sub

oops:
Error_Handler "TileMap1_MouseDown", Err.Description, Err.Number
End Sub


Private Sub TileMap1_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo oops
Dim mX As Integer, mY As Integer
mX = Int(X / Map_Zoom)
mY = Int(Y / Map_Zoom)
Label7 = "Block Coordinate: [" & mX & "-" & mY & "]"
If TileMap1.Tag = "m" Then
    If button = 2 Then
        temp_tool = Selected_Tool + 1
    Else
        temp_tool = Selected_Tool
    End If
    Select Case temp_tool
    Case 0, 9 'pen
        change_map mX, mY, Selected_Block, True
    Case 2 ' brush
        If Shift = 0 Then
            pattern_tool mX, mY
        Else
            Pattern_X = mX: Pattern_Y = mY
            pattern_tool mX, mY
        End If
    Case 3, 4 ' select
        temp_width = ((mX + 1) * Map_Zoom) - Shape3.Left
        temp_height = ((mY + 1) * Map_Zoom) - Shape3.Top
        If temp_width <= 0 Then Shape3.width = 0 Else Shape3.width = temp_width
        If temp_height <= 0 Then Shape3.height = 0 Else Shape3.height = temp_height
    Case 1, 5, 7 ' sample
        sample_tool mX, mY
    Case 8 'erase
        Selected_byte = Mid(Store(CL).Unzipped, mY * 256 + mX + 1)
        change_map mX, mY, Asc(Selected_byte), True
    Case 10 'sprites
        If Selected_Sprite = -1 Then
            Exit Sub
        ElseIf Selected_Sprite = 0 Then
            'code
            'Exit Sub
        End If
        Shape3.Left = Int(X / Map_Zoom * 2) * Map_Zoom / 2
        Shape3.Top = (Int(Y / Map_Zoom * 2)) * Map_Zoom / 2
    Case 12, 13 'mag
        Exit Sub
    End Select
    Redraw_minimap
End If
Exit Sub

oops:
Error_Handler "TileMap1_MouseMove", Err.Description, Err.Number
End Sub

Private Sub TileMap1_MouseUp(button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo oops
Dim mX As Integer, mY As Integer, sx As Integer, sy As Integer, block_type As Integer
mX = Int(X / Map_Zoom)
mY = Int(Y / Map_Zoom)
TileMap1.Tag = vbNullString
If button = 2 Then
    temp_tool = Selected_Tool + 1
Else
    temp_tool = Selected_Tool
End If
Select Case temp_tool
Case 0, 2, 8 'pen/brush/erase
    If temp_tool = 2 And Shift = 0 Then
        pattern_tool mX, mY
    End If
Case 3 ' selected
    Shape3.Visible = False
    If Shape3.width = 1 Or Shape3.height = 1 Then Exit Sub
    For cx = 0 To 7
        For cy = 0 To 7
            Pattern(cx, cy) = 0
        Next cy
    Next cx
    Pattern1.Cls
    For sx = 0 To mX - Int(Shape3.Left / Map_Zoom)
        For sy = 0 To mY - Int(Shape3.Top / Map_Zoom)
            block_type = sample_block_map(sx + Int(Shape3.Left / Map_Zoom), sy + Int(Shape3.Top / Map_Zoom))
            set_pattern_block block_type, sx, sy
        Next sy
    Next sx
    scale_pattern
    Pattern1.Refresh
Case 4 'plat
    For sx = lastX To mX
        For sy = LastY To mY
            change_map sx, sy, Selected_Block
        Next sy
    Next sx
    Shape3.Visible = False
    'Redraw_block_Map lastX, LastY, mX, mY
    'Display_sprites
    'TileMap1.Refresh
Case 6 'fill
    For mX = 0 To 255
        change_map mX, mY, Selected_Block
    Next mX
    'Redraw_block_Map 0, LastY, 255, mY
    'Display_sprites
    'TileMap1.Refresh
Case 10 'sprites
    If Selected_Sprite = -1 Then 'new
        Selected_Sprite = UBound(sprite_Store()) + 1
        ReDim Preserve sprite_Store(Selected_Sprite)
        
        sprite_Store(Selected_Sprite).TypeB = Selected_Sprite_Type Mod 8
        sprite_Store(Selected_Sprite).typeA = (Selected_Sprite_Type - sprite_Store(Selected_Sprite).TypeB) / 8
        sprite_Store(Selected_Sprite).X = Int(X / Map_Zoom * 2) + 2
        sprite_Store(Selected_Sprite).Y = Int(Y / Map_Zoom * 2) + 2
    ElseIf Selected_Sprite = 0 Then 'mario
        Dim s_x As Integer, s_y As Integer
        s_x = X \ 256
        s_y = Y \ 256
        'Put_Header CL, 1, CInt(y - s_y * 256)
        'Put_Header CL, 3, CInt(x - s_x * 256)
        'Put_Header CL, 2, s_y
        'Put_Header CL, 4, s_x
        'Put_Header CL, 8, s_x
        ''Put_Header CL, 7, s_x
        
        ''Save_Sprites
        'sprites_Loaded = False
        'Load_sprites CL
    Else
        sprite_Store(Selected_Sprite).X = Int(X / Map_Zoom * 2) + 2
        sprite_Store(Selected_Sprite).Y = Int(Y / Map_Zoom * 2) + 2
    End If
    Shape3.Visible = False
    Display_sprites
Case 11 'sprite delete
    If Selected_Sprite < 0 Then
        Exit Sub
    Else
        Redraw_block_Map mX - 1, mY - 1, mX + 1, mY + 1
        sprite_Store(Selected_Sprite).X = 514
        Shape3.Visible = False
        Display_sprites
    End If
Case 12, 13 'mag
    Exit Sub
Case 1, 5, 7 'sample
    sample_tool mX, mY
End Select
Display_Layers
'DoEvents
'Redraw_minimap
Exit Sub

oops:
Error_Handler "TileMap1_MouseUp", Err.Description, Err.Number
End Sub

Private Sub Redraw_minimap()
SetStretchBltMode MiniMap1.hdc, HALFTONE
StretchBlt MiniMap1.hdc, 0, 0, 256, 48, LayerDC(1), 0, 0, 256 * 16, 48 * 16, SRCCOPY
MiniMap1.Refresh
End Sub

Private Sub TileMap1_Resize()
ScrollBars
End Sub

Private Sub change_tool(ByVal New_Tool As Integer)
ToolBox1(Selected_Tool).Enabled = True
ToolBox1(New_Tool).Enabled = False
ToolBox1(Selected_Tool + 1).Enabled = True
ToolBox1(New_Tool + 1).Enabled = False

Selected_Tool = New_Tool
ToolBox1_Click New_Tool
End Sub

Private Sub sample_tool(mX As Integer, mY As Integer)
    Selected_Block = sample_block_map(mX, mY)
    block_y = Int(Selected_Block / 16)
    block_x = Selected_Block - block_y * 16
    Shape1.Left = block_x * 16
    Shape1.Top = block_y * 16
End Sub

Private Function sample_block_map(mX As Integer, mY As Integer)
    If mX < 0 Or mY < 0 Then Exit Function
    sample_block_map = Asc(Mid$(Store(CL).Unzipped_Modified, mY * 256 + mX + 1, 1))
End Function

Private Sub change_map(mX As Integer, mY As Integer, block_type As Integer, Optional Draw As Boolean)
If Store(CL).Unzipped_Modified = vbNullString Then Exit Sub
If mX < 0 Or mY < 0 Or mX > 255 Or mY > 47 Then Exit Sub
block_y = Int(block_type / 16)
block_x = block_type - block_y * 16
If Draw Then
    StretchBlt TileMap1.hdc, mX * Map_Zoom, mY * Map_Zoom, Map_Zoom, Map_Zoom, Blocks1.hdc, block_x * 16, block_y * 16, 16, 16, SRCCOPY
    TileMap1.Refresh
End If
BitBlt LayerDC(1), mX * 16, mY * 16, 16, 16, Blocks1.hdc, block_x * 16, block_y * 16, SRCCOPY

Mid(Store(CL).Unzipped_Modified, mY * 256 + mX + 1) = Chr$(block_type)
End Sub

Private Sub ScrollBars()
HScroll1.Max = -(Picture2.ScaleWidth - TileMap1.width)
VScroll1.Max = -(Picture2.ScaleHeight - TileMap1.height)
If TileMap1.ScaleWidth < Picture2.ScaleWidth Then
    TileMap1.Left = -1
    HScroll1.Visible = False
Else
    HScroll1.Visible = True
End If
If TileMap1.ScaleHeight < Picture2.ScaleHeight Then
    TileMap1.Top = -1
    VScroll1.Visible = False
Else
    VScroll1.Visible = True
    If Headers_Start < 0 Then
        VScroll1.value = VScroll1.Max / 2
    End If
End If

Display_ScreenPos
End Sub

Private Sub TileMap2_mouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_TileMap2 0, button, X, Y
End Sub

Private Sub tilemap2_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_TileMap2 1, button, X, Y
End Sub

Private Sub tilemap2_MouseUp(button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_TileMap2 2, button, X, Y
End Sub

Private Sub Mouse_TileMap2(action As Integer, button As Integer, X As Single, Y As Single)
Dim Selected_Tile As Integer, Tile_X As Integer, Tile_Y As Integer, v As String
    
Tile_X = Int(X / 8)
Tile_Y = Int(Y / 8)

If button = 2 Then
    If X < 0 Or Y < 0 Or X > tilemap2.ScaleWidth - 1 Or Y > tilemap2.ScaleHeight - 1 Then Exit Sub
    
    TX = Tile_X + px: TY = Tile_Y + py
            
    If Tilemap_Mode = "title" Then
        TX = TX + 7
    ElseIf Tilemap_Mode = "4" Then
        If TX >= 32 And TY >= 32 Then
            TY = TY + 63
        ElseIf TX >= 32 Then
            TY = TY + 31
        ElseIf TY >= 32 Then
            TY = TY + 32
        End If
    End If
    Selected_Tile = Asc(Mid(ROM_Data, tilemap_start + TX + TY * 32 + 1, 1))
    
    Selected_Tile = (Selected_Tile + 128) Mod 256
    
    py = Selected_Tile \ 16
    px = Selected_Tile - py * 16
    Shape7.Top = py * 8
    Shape7.Left = px * 8
    'Selected_Tile = (Shape7.Top \ 8 + py) * 16 + (Shape7.Left \ 8 + px) + 128
    'Selected_Tile = Selected_Tile Mod 256
End If
Select Case action
Case 0
    tilemap2.Tag = "m"
Case 2
    tilemap2.Tag = vbNullString
End Select


If tilemap2.Tag = "m" Then

    If X < 0 Or Y < 0 Or X > tilemap2.ScaleWidth - 1 Or Y > tilemap2.ScaleHeight - 1 Then Exit Sub



    BitBlt tilemap2.hdc, Tile_X * 8, Tile_Y * 8, Shape7.width, Shape7.height, tiles2.hdc, Shape7.Left, Shape7.Top, SRCCOPY
    tilemap2.Refresh
    DoEvents
    
    For px = 0 To Shape7.width \ 8 - 1
        For py = 0 To Shape7.height \ 8 - 1
            Selected_Tile = (Shape7.Top \ 8 + py) * 16 + (Shape7.Left \ 8 + px) + 128
            Selected_Tile = Selected_Tile Mod 256
            
            TX = Tile_X + px: TY = Tile_Y + py
            
            If Tilemap_Mode = "title" Then
                TX = TX + 7
            ElseIf Tilemap_Mode = "4" Then
                If TX >= 32 And TY >= 32 Then
                    TY = TY + 63
                ElseIf TX >= 32 Then
                    TY = TY + 31
                ElseIf TY >= 32 Then
                    TY = TY + 32
                End If
            End If
            Mid(ROM_Data, tilemap_start + TX + TY * 32 + 1, 1) = Chr(Selected_Tile)
        Next py
    Next px
End If
End Sub
Private Sub tilemap2_Resize()
xover = tilemap2.width - Picture4.ScaleWidth - 2
yover = tilemap2.height - Picture4.ScaleHeight - 2
If yover > 0 Then
    tilemap_VScroll.Visible = True
    tilemap_VScroll.Max = yover
Else
    tilemap_VScroll.Visible = False
End If
If xover > 0 Then
    tilemap_HScroll.Visible = True
    tilemap_HScroll.Max = xover
Else
    tilemap_HScroll.Visible = False
End If

End Sub

Private Sub Tiles1_MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Tiles button, X, Y, 1
End Sub

Private Sub Mouse_Tiles(button As Integer, X As Single, Y As Single, Down As Integer)
sx = Int(X / 8)
sy = Int(Y / 8)

If Down = 1 Then
    Shape2.Tag = "m"
ElseIf Down = 3 Then
    Shape2.Tag = vbNullString
End If

If X < 0 Or Y < 0 Or X > Tiles1.ScaleWidth - 1 Or Y > Tiles1.ScaleHeight - 1 Then Exit Sub

If Shape2.Tag = "m" Then
    Shape2.Left = sx * 8
    Shape2.Top = sy * 8
End If

End Sub

Private Sub Tiles1_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Tiles button, X, Y, 2
End Sub

Private Sub Tiles1_MouseUp(button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Tiles button, X, Y, 3
End Sub

Private Sub tiles2_MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Tiles2 0, button, X, Y
End Sub

Private Sub tiles2_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Tiles2 1, button, X, Y
End Sub

Private Sub tiles2_MouseUp(button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Tiles2 2, button, X, Y
End Sub

Private Sub Mouse_Tiles2(action As Integer, button As Integer, X As Single, Y As Single)
If X < 0 Or Y < 0 Or X > tiles2.ScaleWidth - 1 Or Y > tiles2.ScaleHeight - 1 Then Exit Sub
Select Case action
Case 0
    Shape7.Left = (X \ 8) * 8
    Shape7.Top = (Y \ 8) * 8
    Shape7.height = 8
    Shape7.width = 8
    
    Shape7.Tag = "m"
Case 1, 2
    If Shape7.Tag = "m" Then
        If X < Shape7.Left Then X = Shape7.Left
        If Y < Shape7.Top Then Y = Shape7.Top
        sx = X \ 8
        sy = Y \ 8
        Shape7.width = sx * 8 - Shape7.Left + 8
        Shape7.height = sy * 8 - Shape7.Top + 8
    End If
    If action = 2 Then Shape7.Tag = vbNullString
End Select

End Sub

Private Sub Timer1_Timer()
If Shape1.drawmode = 7 Then
    Shape1.drawmode = 10
    Shape2.drawmode = 10
    Shape3.drawmode = 10
    Shape5.drawmode = 10
    Shape6.drawmode = 10
    Shape7.drawmode = 10
Else
    Shape1.drawmode = 7
    Shape2.drawmode = 7
    Shape3.drawmode = 7
    Shape5.drawmode = 7
    Shape6.drawmode = 7
    Shape7.drawmode = 7
End If
End Sub

Private Sub ToolBox1_Click(Index As Integer)
If Index Mod 2 = 1 Then
    Index = Index - 1
End If

If Index <> Selected_Tool Then change_tool Index
Picture3.Visible = False
Blocks2.Visible = False
Pattern1.Visible = False
VScroll3.Visible = False
Select Case Index
Case 0, 4, 6 'pen,plat,fill,pattern
    Blocks2.Visible = True
Case 2 'brush
    Blocks2.Visible = True
    Pattern1.Visible = True
Case 10 'sprites
    Picture3.Visible = True
    VScroll3.Visible = True
End Select
Form_Resize
End Sub

Private Sub VScroll1_Change()
TileMap1.Top = -VScroll1.value - 1
Display_ScreenPos
End Sub

Public Sub Load_ROM(ROM As String)
Dim True_Rom_Version As String
On Error GoTo oops
Dim temp As String
If ROM = vbNullString Then
    Lock_GUI
    Exit Sub
End If
Unload_ROM
ROM_Data = Space(FileLen(ROM))
Close 1
Open ROM For Binary As 1
Get #1, , ROM_Data

temp = Mid(ROM_Data, 309, 19)
ROM_Name = Left(temp, InStr(1, temp, Chr(0)) - 1)

True_Rom_Version = Asc(Mid(ROM_Data, 333, 1))
Text3 = True_Rom_Version
If True_Rom_Version > 2 Then Rom_Version = 2 Else Rom_Version = True_Rom_Version

Lock_GUI True

Text2 = ROM_Name
Select Case LCase(ROM_Name)
'Case "mario 1", "super marioland", "mario land 1"
    'ROM_Name = "Mario 1"
Case "mario 2", "marioland2", "mario 2 mod", "mario mand 2", "mario 2b"
    Status "Loaded ROM: " & ROM_Name & " - Totally supported", 2
    Select Case Rom_Version
    Case 0, 1, 2, 3
        If LCase(ROM_Name) = "mario 2 mod" Then
            Label11 = "Modified SML2 ROM (v1." & Rom_Version & ")"
            Label11.ForeColor = vbRed
        Else
            Label11 = "SML2 ROM v1." & Rom_Version
            Label11.ForeColor = vbBlack
        End If
    Case 4
        Label11.ForeColor = &H80FF&
        Label11 = "SML Mod - Beta release"
    Case Else
        Label11.ForeColor = vbBlue
        Label11 = "SML 2" & Chr(Asc("a") + (Rom_Version \ 10)) & " (v" & Rom_Version Mod 10 & ")"
    End Select
    ROM_Name = "Mario 2"
Case "wario", "supermarioland3", "mario land 3", "mario 3b"
    Label11 = "SWL ROM v1." & Rom_Version
    Label11.ForeColor = &H80FF&
    Status "Loaded ROM: 'Wario 1'", 2
    Status "No header support", 3
    ROM_Name = "Wario"
    SSTab2.TabEnabled(1) = False
    Check8(3).value = False ': Check8(3).Enabled = False
Case Else
    Status "ROM not supported: '" & ROM_Name & "'", 4
    GoTo oops
End Select

Load_Headers
Load_Config
If Headers_Start > 0 Then Load_All_Sprites
If ROM_Name = "Mario 2" Then Load_All_Borders
Load_Zone_List

Exit Sub
oops:
Lock_GUI
If Err.Number Then Error_Handler "Loading ROM", Err.Description, Err.Number, ROM_Name
Unload_ROM
WritePrivateProfileString "config", "rom", "", App.Path & "\config\config.ini"
Status "Clearing AutoLoad", 3
End Sub

Private Sub Lock_GUI(Optional Enabled As Boolean)
If Enabled = False Then
    SSTab2.Tab = 3
    Patching.Visible = False
End If

SSTab2.TabEnabled(0) = Enabled
SSTab2.TabEnabled(1) = Enabled
SSTab2.TabEnabled(2) = Enabled
menu_edit.Enabled = Enabled
Menu_Patch_ROM = Enabled
menu_export = Enabled
menu_import = Enabled
menu_save_images = Enabled
End Sub
Private Sub Load_Headers()
On Error GoTo oops
Dim Level As Integer, temp As String
Headers_Start = From_ini("config", "headers", 22028, True)
ReDim Bank_Store(0)
'1=y 2=sy 3=x 4=sx 5=sy  7=scroll 8=castle 11=level 12=zero 13=block 14=bank 16=pallette 19=subbank 20=time
Do
    ReDim Preserve Store(Level)
    If Headers_Start > 0 Then
        SSTab1.TabEnabled(3) = True
        If Get_Header(Level, 12) > 0 Or Get_Header(Level, 11) <> Level Then
            Store(Level).Map_Bank = -1
            Exit Sub
        End If
        
        Store(Level).Block_Bank = Get_Header(Level, 13)
        Store(Level).Map_Bank = Get_Header(Level, 14)
        Store(Level).Map_Sub_Bank = Get_Header(Level, 19)
    Else
        SSTab1.TabEnabled(3) = False
        If Level >= From_ini("config", "levels", 43, True) Then Exit Sub
        temp = From_ini("maps", Level & "bank", "0,0")
        Store(Level).Map_Bank = Split(temp, ",")(0)
        Store(Level).Map_Sub_Bank = Split(temp, ",")(1)
    End If
    Add_Bank Store(Level).Map_Bank, Store(Level).Map_Sub_Bank, Level
    Level = Level + 1
Loop
Exit Sub

oops:
Error_Handler "Load_Headers", Err.Description, Err.Number

End Sub

Private Sub Add_Bank(Bank As Integer, Sub_Bank As Integer, Level As Integer)
If UBound(Bank_Store()) > 0 Then
    B = -1
    Do
    
    B = B + 1
    Loop Until Bank_Store(B).Bank = Bank Or B >= UBound(Bank_Store())
Else
    B = 0
End If
If Bank_Store(B).Bank <> Bank Then
    If Bank_Store(B).Bank > 0 Then B = B + 1
    ReDim Preserve Bank_Store(B)
    Bank_Store(B).Bank = Bank
End If
If Sub_Bank > Bank_Store(B).Last_Sub Then
    Bank_Store(B).Last_Sub = Sub_Bank
End If
If Bank_Store(B).Level_list <> vbNullString Then spacer = ","
Bank_Store(B).Level_list = Bank_Store(B).Level_list & spacer & Level
End Sub
Public Sub Unload_ROM()
ReDim Store(0)
ReDim sprite_Store(0)
ReDim Bank_Store(0)
CL = -1
sprites_Loaded = False
ROM_Name = vbNullString: ROM_Data = vbNullString
Combo1.Clear: Combo2.Clear
Combo1.Text = "Zone": Combo2.Text = "Map"
Combo1.Enabled = False: Combo2.Enabled = False
TileMap1.Cls: TileMap1.width = 0: TileMap1.height = 0
MiniMap1.Visible = False
Label7 = vbNullString
End Sub

Private Sub Load_Level(Level As Integer, Optional Redraw_Mode As Integer)
'On Error GoTo oops:
Dim Sprites_From As Long

If CL <> Level Or Level = -1 Then
    If Level = -1 Then Exit Sub
    If sprites_Loaded Then
        Save_Sprites
        sprites_Loaded = False
    End If
    CL = Level
    
    redo_blocks = True
    If Headers_Start > 0 Then
        Update_Palette
    Else
        set_palette palette1(0).BackColor, palette1(1).BackColor, palette1(2).BackColor, palette1(3).BackColor
    End If
ElseIf Store(CL).Map_Bank = -1 Then
    Status "Invalid Header", 4
    Beep
    Exit Sub
End If

If Combo1.ListIndex = -1 Then Exit Sub

Display_Headers

If Headers_Start > 0 Then
    Status "Loading: Level " & Level & " [" & Store(CL).Map_Bank & ":" & Store(CL).Map_Sub_Bank & "]", 2
  
    If Redraw_Mode > 0 Then
        Load_Tiles Level, Get_Tile_Start(Level), 125441
        Load_Blocks CLng(Store(Level).Block_Bank), 114688
        If Redraw_Mode = 2 Then Exit Sub
    End If
    
    Display_block_Map Level
    
    Redraw_minimap
    
    If Check1(0) Then Load_sprites Level
    
    If Check8(3) Then Display_Borders Level
Else
    If Redraw_Mode > 0 Then
        Load_Tiles 0, Get_Tile_Start(Level), 282609
        Load_Blocks CLng(Level), 181770
        If Redraw_Mode = 2 Then Exit Sub
    End If

    Check1(0).value = False
    
    Display_block_Map Level
    
    Redraw_minimap
End If

If MiniMap1.ScaleHeight > 1 Then MiniMap1.Visible = True

Display_Layers

'TileMap1.Visible = True
Exit Sub

oops:
Display_Locked = False
Error_Handler "Load Level", Err.Description, Err.Number
End Sub

Private Sub Display_Borders(Level)
Dim stops As Integer
ColourBuffer LayerDC(3), 256 * 16, 48 * 16, vbWhite
BluePen = CreatePen(PS_SOLID, 1, vbBlue)

For X = 0 To 15
    For Y = 0 To 2
        stops = Asc(Mid(Store(Level).Borders, Y * 16 + X + 1, 1))
        If stops >= 8 Then 'bottom
            stops = stops - 8
            BitBlt LayerDC(3), (X * 16 + 8) * 16, (Y * 16 + 15) * 16, 16, 16, special_blocks1.hdc, 272, 0, SRCCOPY
            DrawGDILine LayerDC(3), ((X) * 16) * 16, ((Y + 1) * 16) * 16, ((X + 1) * 16) * 16, ((Y + 1) * 16) * 16
        End If
        If stops >= 4 Then 'top
            stops = stops - 4
            BitBlt LayerDC(3), (X * 16 + 8) * 16, (Y * 16) * 16, 16, 16, special_blocks1.hdc, 272, 16, SRCCOPY
            DrawGDILine LayerDC(3), ((X) * 16) * 16, ((Y) * 16) * 16, ((X + 1) * 16) * 16, ((Y) * 16) * 16
        End If
        If stops >= 2 Then 'left
            stops = stops - 2
            BitBlt LayerDC(3), (X * 16) * 16, (Y * 16 + 8) * 16, 16, 16, special_blocks1.hdc, 272, 32, SRCCOPY
            DrawGDILine LayerDC(3), ((X) * 16) * 16, ((Y) * 16) * 16, ((X) * 16) * 16, ((Y + 1) * 16) * 16
        End If
        If stops >= 1 Then 'right
            stops = stops - 1
            BitBlt LayerDC(3), (X * 16 + 15) * 16, (Y * 16 + 8) * 16, 16, 16, special_blocks1.hdc, 272, 48, SRCCOPY
            DrawGDILine LayerDC(3), ((X + 1) * 16) * 16, ((Y) * 16) * 16, ((X + 1) * 16) * 16, ((Y + 1) * 16) * 16
        End If
        If stops > 0 Then Status "Weird border data", 3
Next Y, X

For X = 0 To 15
    For Y = 0 To 2
        If Asc(Mid(Store(Level).Warps, Y * 16 + X + 1, 1)) > 0 Then
            RedPen = SelectObject(LayerDC(3), BluePen)
            DrawGDILine LayerDC(3), ((X) * 16) * 16, ((Y + 1) * 16) * 16, ((X + 1) * 16) * 16, ((Y + 1) * 16) * 16
            DrawGDILine LayerDC(3), ((X) * 16) * 16, ((Y) * 16) * 16, ((X + 1) * 16) * 16, ((Y) * 16) * 16
            DrawGDILine LayerDC(3), ((X) * 16) * 16, ((Y) * 16) * 16, ((X) * 16) * 16, ((Y + 1) * 16) * 16
            DrawGDILine LayerDC(3), ((X + 1) * 16) * 16, ((Y) * 16) * 16, ((X + 1) * 16) * 16, ((Y + 1) * 16) * 16
            BluePen = SelectObject(LayerDC(3), RedPen)
        End If
Next Y, X
End Sub


Private Sub Display_Layers()
TileMap1.width = 256 * Map_Zoom + 2
TileMap1.height = 48 * Map_Zoom + 2
DoEvents

BitBlt LayerDC(0), 0, 0, 256 * 16, 48 * 16, LayerDC(1), 0, 0, vbSrcCopy
If Check1(0) Then TransBltOverlay LayerDC(0), LayerDC(2), 0, 0, 256 * 16, 48 * 16, vbWhite
If Check8(3) Then TransBltOverlay LayerDC(0), LayerDC(3), 0, 0, 256 * 16, 48 * 16, vbWhite

'BitBlt TileMap1.hdc, 0, 0, 256 * 16, 48 * 16, LayerDC(3), 0, 0, SRCCOPY

If Check1(2) = 1 Then mode = HALFTONE Else mode = 3

SetStretchBltMode TileMap1.hdc, mode
StretchBlt TileMap1.hdc, 0, 0, 256 * Map_Zoom, 48 * Map_Zoom, LayerDC(0), 0, 0, 256 * 16, 48 * 16, SRCCOPY
TileMap1.Visible = True
TileMap1.Refresh
DoEvents
End Sub
Private Sub Display_Headers()
Text1(2) = Store(CL).Map_Bank
Text1(3) = Store(CL).Map_Sub_Bank
Text1(4) = Store(CL).Block_Bank
Text1(0) = Get_Header(CL, 15)
Text1(1) = 100 * Get_Header(CL, 20)
If Get_Header(CL, 7) = 80 Then
    Check10.value = 0
Else
    Check10.value = 1
End If
End Sub
Public Sub Error_Handler(Location As String, desc As String, err_num As Integer, Optional More_desc As String)
If SSTab2.Visible = True Then SSTab2.Tab = 3
If More_desc <> vbNullString Then
    More_desc = " More info: '" & More_desc & "'"
End If
Status desc & " [" & err_num & "] Caught in: '" & Location & "'" & More_desc, 4
Beep
End Sub

Private Sub Auto_Center()
Dim X As Integer, Y As Integer
mY = Get_Header(CL, 1)
mX = Get_Header(CL, 3)
sy = Get_Header(CL, 2)
sx = Get_Header(CL, 4)
map_scale = Map_Zoom / 16

Center_On CLng(mX / 16 + (sx + 1) * 16 + 1), CLng(sy * 16 + Picture2.ScaleHeight / 2 / 16 + 1)
End Sub

Private Sub Center_On(X As Integer, Y As Integer)
Dim TempX As Single, TempY As Single
map_scale = Map_Zoom / 16

If VScroll1.Max > 0 Then
    
    TempY = Y * map_scale * 16 - Picture2.ScaleHeight / 2
    
    If TempY < 0 Then
        VScroll1.value = 0
    ElseIf TempY < VScroll1.Max Then
        VScroll1.value = TempY
    Else
        VScroll1.value = VScroll1.Max
    End If

End If

If HScroll1.Max > 0 Then
    
    TempX = X * map_scale * 16 - Picture2.ScaleWidth / 2
    
    If TempX < 0 Then
        HScroll1.value = 0
    ElseIf TempX < HScroll1.Max Then
        HScroll1.value = TempX
    Else
        HScroll1.value = HScroll1.Max
    End If
End If

Display_ScreenPos

End Sub

Private Sub Display_ScreenPos()
Shape4.Left = HScroll1.value / Map_Zoom
Shape4.Top = VScroll1.value / Map_Zoom
Shape4.width = Picture2.ScaleWidth / Map_Zoom
Shape4.height = Picture2.ScaleHeight / Map_Zoom
If Shape4.width > MiniMap1.ScaleWidth Then Shape4.width = MiniMap1.ScaleWidth
If Shape4.height > MiniMap1.ScaleHeight Then Shape4.height = MiniMap1.ScaleHeight

End Sub

Private Sub Update_Palette()
Dim temp(3) As Integer, Palette_Code As Integer
Palette_Code = Get_Header(CL, 16)

temp(0) = Palette_Code And 3
temp(1) = (Palette_Code And 12) / 4
temp(2) = (Palette_Code And 48) / 16
temp(3) = (Palette_Code And 192) / 64

set_palette palette1(3 - temp(3)).BackColor, palette1(3 - temp(2)).BackColor, palette1(3 - temp(1)).BackColor, palette1(3 - temp(0)).BackColor
End Sub
Private Sub Display_Screen_Border(mY As Integer, sy As Integer)
    TileMap1.Line (0, TileMap1.ScaleHeight - mY)-(22 * 16, TileMap1.ScaleHeight - sy), RGB(0, 255, 0) ', B
End Sub
Private Function Skip_To_sprites(sprites_start As Long, Level As Integer) As Long
Dim B As Long, found_level As Integer
B = sprites_start
Do
    If Mid$(ROM_Data, B + 1, 1) = "ÿ" Then
        found_level = found_level + 1
        B = B + 1
    End If
    B = B + 3
Loop Until found_level = Level
B = B - 3
Skip_To_sprites = B
End Function

Private Sub Load_All_Sprites()
Dim Level As Integer
If Headers_Start = 0 Then Exit Sub
Level = -1
Do
    Level = Level + 1
    Sprites_From = Skip_To_sprites(57463, Level)
    Sprites_to = Sprites_From
    Do
        Sprites_to = InStr(Sprites_to + 2, ROM_Data, "ÿ")
        'z = (Sprites_to - Sprites_From) Mod 3
    Loop Until (Sprites_to - Sprites_From) Mod 3 = 1
    total_sprites = Int((Sprites_to - Sprites_From) / 3)
    Store(Level).Sprites = Mid$(ROM_Data, Sprites_From + 1, total_sprites * 3)
Loop Until Get_Header(Level + 1, 12) > 0
End Sub

Private Sub Load_All_Borders()
Dim Level As Integer
Dim ptr As Long
If Headers_Start = 0 Then Exit Sub
Level = -1
Do
    Level = Level + 1

    ptr = CLng(16384) * Store(Level).Map_Bank
    ptr = ptr + 5 + Store(Level).Map_Sub_Bank * 48
    Store(Level).Borders = Mid(ROM_Data, ptr, 48)
    For B = 0 To UBound(Bank_Store)
        If Bank_Store(B).Bank = Store(Level).Map_Bank Then Exit For
    Next B
    ptr = ptr + (Bank_Store(B).Last_Sub + 1) * 48
    Store(Level).Warps = Mid(ROM_Data, ptr, 48)
Loop Until Get_Header(Level + 1, 12) > 0
End Sub

Private Sub Save_ALl_Borders()
Dim Level As Integer
Dim ptr As Long
If Headers_Start = 0 Then Exit Sub
Level = -1
Do
    Level = Level + 1

    ptr = CLng(16384) * Store(Level).Map_Bank
    ptr = ptr + Asc(Mid(ROM_Data, ptr + 1, 1)) + Store(Level).Map_Sub_Bank * 48
    Mid(ROM_Data, ptr + 1, 48) = Store(Level).Borders
Loop Until Get_Header(Level + 1, 12) > 0
End Sub
Private Sub Load_sprites(Level As Integer)
Dim pos As Long, Sprites_From As Long
If sprites_Loaded Then GoTo skip_load

If Store(Level).Sprites = vbNullString Then
    Sprites_From = Skip_To_sprites(57463, Level)
    Sprites_to = Sprites_From
    Do
        Sprites_to = InStr(Sprites_to + 2, ROM_Data, "ÿ")
        'z = (Sprites_to - Sprites_From) Mod 3
    Loop Until (Sprites_to - Sprites_From) Mod 3 = 1
    total_sprites = Int((Sprites_to - Sprites_From) / 3)
    Store(Level).Sprites = Mid$(ROM_Data, Sprites_From + 1, total_sprites * 3)
End If
Load_sprite_aliases Level

Selected_Sprite_Type = 1
change_selected_sprite

mY = Get_Header(Level, 1)
mX = Get_Header(Level, 3)
sy = Get_Header(Level, 2)
sx = Get_Header(Level, 4)

sprite_Store(0).X = (mX + sx * 256)
sprite_Store(0).Y = (mY + sy * 256)

Do
    ReDim Preserve sprite_Store(pos + 1)
    A = Asc(Mid$(Store(Level).Sprites, pos * 3 + 1, 1))
    B = Asc(Mid$(Store(Level).Sprites, pos * 3 + 2, 1))
    c = Asc(Mid$(Store(Level).Sprites, pos * 3 + 3, 1))
    x1 = A And 15
    x2 = B And 31
    sprite_Store(pos + 1).Y = (c And 127)
    sprite_Store(pos + 1).X = (x1 * 32 + x2 + 1)

    sprite_Store(pos + 1).typeA = (A And 240) / 16
    sprite_Store(pos + 1).TypeB = (B And 224) / 32
    pos = pos + 1
Loop Until (pos + 1) * 3 > Len(Store(Level).Sprites) 'Mid$(ROM_Data, Sprites_From + pos * 3 + 1, 1) = "ÿ"
sprites_Loaded = True
skip_load:
Display_sprites
End Sub
Private Sub change_selected_sprite()
X = Selected_Sprite_Type Mod 8
Y = (Selected_Sprite_Type - X) / 8
Shape5.Left = X * 32
Shape5.Top = Y * 32
End Sub
Private Sub Display_sprites()
Status "Displaying: Sprites", 1

ColourBuffer LayerDC(2), 256 * 16, 48 * 16, vbWhite

map_scale = Map_Zoom / 16
'TransparentBlt TileMap1.hDC, sprite_Store(0).X * map_scale, sprite_Store(0).Y * map_scale, Map_Zoom * 2, Map_Zoom * 2, sprites2.hDC, 0, 0, 32, 32, RGB(255, 255, 255)
BitBlt LayerDC(2), sprite_Store(0).X, sprite_Store(0).Y, 32, 32, sprites2.hdc, 0, 0, SRCCOPY

For pos = 1 To UBound(sprite_Store())
    'TransparentBlt TileMap1.hDC, (sprite_Store(pos).X / 2 - 1) * Map_Zoom, (sprite_Store(pos).Y / 2 - 2) * Map_Zoom, Map_Zoom * 2, Map_Zoom * 2, sprites2.hDC, sprite_Store(pos).TypeB * 32, sprite_Store(pos).typeA * 32, 32, 32, RGB(255, 255, 255)
    BitBlt LayerDC(2), (sprite_Store(pos).X / 2 - 1) * 16, (sprite_Store(pos).Y / 2 - 2) * 16, 16 * 2, 16 * 2, sprites2.hdc, sprite_Store(pos).TypeB * 32, sprite_Store(pos).typeA * 32, SRCCOPY
Next pos
End Sub

Private Sub Save_Sprites()
Dim Sprites As Long
If sprites_Loaded = False Then Exit Sub
Status "Saving: sprites", 2
Order_Sprites
'Store(CL).Sprites = String(UBound(sprite_Store()) * 3 + 1, "#")
Store(CL).Sprites = vbNullString
''Sprites_From = Skip_To_sprites(57463, CL)

For pos = 0 To UBound(sprite_Store()) - 1
    If sprite_Store(pos + 1).X > 512 Then
        GoTo done
    End If
    x2 = (sprite_Store(pos + 1).X - 1) And 31
    x1 = ((sprite_Store(pos + 1).X - x2) / 32) And 15
    
    
    A = ((sprite_Store(pos + 1).typeA * 16) And 240) Or x1
    B = ((sprite_Store(pos + 1).TypeB * 32) And 244) Or x2
    c = sprite_Store(pos + 1).Y And 127
    
    'Mid$(Store(CL).Sprites, pos * 3 + 1, 1) = Chr(a)
    'Mid$(Store(CL).Sprites, pos * 3 + 2, 1) = Chr(b)
    'Mid$(Store(CL).Sprites, pos * 3 + 3, 1) = Chr(c)
    Store(CL).Sprites = Store(CL).Sprites & Chr(A) & Chr(B) & Chr(c)
Next pos
done:
'Mid$(Store(CL).Sprites, pos * 3 + 1, 1) = Chr(255)
''Store(CL).Sprites = Left$(Store(CL).Sprites, pos * 3 + 1)
''Mid$(ROM_Data, Sprites_From + 1) = Store(CL).Sprites
End Sub
Private Sub Order_Sprites()
Dim temp As Sprite_Store_type
Do
wrong = 0
For pos = 1 To UBound(sprite_Store()) - 1
    If sprite_Store(pos).X > sprite_Store(pos + 1).X Then
        wrong = wrong + 1
        temp = sprite_Store(pos)
        sprite_Store(pos) = sprite_Store(pos + 1)
        sprite_Store(pos + 1) = temp
    End If
Next pos
Loop Until wrong = 0
End Sub

Private Sub Load_sprite_aliases(Level As Integer)
Dim alias_rule As String
sprites1 = LoadPicture(App.Path & "\config\" & From_ini("config", "sprites")) 'App.Path & "/config/" & ROM_Name & " - sprites.gif")
BitBlt sprites2.hdc, 0, 0, 256, 512, sprites1.hdc, 0, 0, SRCCOPY
A = 1
Do Until From_ini("sprites", Level & "-" & A, -1, True) = -1
    alias_rule = From_ini("sprites", Level & "-" & A, "-1")
    If alias_rule = "-1" Then Exit Sub
    
    before = Split(alias_rule, ">")(0)
    beforey = Val(Split(before, ",")(0))
    beforex = Val(Split(before, ",")(1))
    If UBound(Split(alias_rule, ">")) = 0 Then
        aftery = beforey
        afterx = beforex + 8
    Else
        after = Split(alias_rule, ">")(1)
        aftery = Val(Split(after, ",")(0))
        afterx = Val(Split(after, ",")(1))
    End If
    
    BitBlt sprites2.hdc, 32 * beforex, 32 * beforey, 32, 32, sprites1.hdc, 32 * afterx, 32 * aftery, SRCCOPY
    A = A + 1
Loop
'Next a
BitBlt sprites3.hdc, 0, 0, 256, 512, sprites2.hdc, 0, 0, SRCCOPY
DoEvents
sprites2.Refresh
sprites3.Refresh

End Sub
Private Function Get_Header(Level As Integer, header As Integer)
'1=my 2=msy 3=mx 4=msx
'5=sy 6=spy 7=sfocusx, 8=focusx,
'9?=castle 10?=x/scroll 11=level 12=zero 13=block 14=bank 16=pallette 19=subbank 20=time
If Headers_Start <= 0 Then Exit Function
Get_Header = Asc(Mid(ROM_Data, Headers_Start + Level * 20 + header - 1, 1))
End Function

Private Function Put_Header(Level As Integer, header As Integer, value As Integer)
If Headers_Start <= 0 Then Exit Function
Mid(ROM_Data, Headers_Start + Level * 20 + header - 1, 1) = Chr(value)
End Function

Public Sub Status(Info As String, Optional MSGlevel As Integer, Optional SameLine As Boolean, Optional patch As Boolean)
Dim Warninglevel As Integer, TypeText As String, NextLine As String, typeflag As String
'Warninglevel = Combo4.ListIndex
If Warninglevel = -1 Then Warninglevel = 0

'If Warninglevel <= MSGlevel Then
If SameLine Then
    Status_data = Left$(Status_data, Len(Status_data) - 1) & Info & "|"
    Refresh_Status patch, False
Else
    If MSGlevel = 3 Then
        TypeText = "*Warning* - "
    ElseIf MSGlevel = 4 Then
        TypeText = "*ERROR* - "
    End If
    'If Len(Status_data) = 0 Then NextLine = "" Else NextLine = vbCrLf
    If patch Then typeflag = "p" Else typeflag = "s"
    Status_data = Status_data & typeflag & "~" & MSGlevel & "~" & Time & "~" & TypeText & Info & "|"
    Refresh_Status patch, True
End If
'End If
End Sub

Public Sub Refresh_Status(Optional patch As Boolean, Optional last_only As Boolean)
Dim s As Integer, e As Integer, temp() As String, temp2() As String
If Status_data = vbNullString Then Exit Sub
If last_only Then
    last_line = InStrRev(Status_data, "|", Len(Status_data) - 1)
    ReDim temp(0)
    temp(0) = Mid(Status_data, last_line + 1, Len(Status_data) - last_line - 1)
    'temp(0) = Replace(temp1, vbCrLf, vbNullString)
    s = 0: e = 0
Else
    Patching.status_text = vbNullString
    status_text = vbNullString
    
    temp() = Split(Status_data, "|")
    s = 0: e = UBound(temp())
End If
For L = s To e
    If Len(temp(L)) > 2 Then
        temp2() = Split(temp(L), "~")
        
        If temp2(0) = "p" And Patching.Combo1.ListIndex <= Val(temp2(1)) Then
            Patching.status_text = Patching.status_text & temp2(3) & vbCrLf
        End If
        If Combo4.ListIndex <= Val(temp2(1)) Then
            status_text = status_text & "[" & temp2(2) & "] - " & temp2(3) & vbCrLf
        End If
    End If
Next L
status_text.SelStart = Len(status_text)
If patch Then Patching.status_text.SelStart = Len(Patching.status_text)
End Sub

Private Function Skip_To(Map_Bank As Integer, Sub_Bank As Integer)
Dim Start_from As Long

If ROM_Name = "Mario 2" Then offset = 1281 Else offset = 65

Start_from = CLng(Map_Bank) * 16384 + offset

If Sub_Bank = 0 Then
    Skip_To = Start_from
    Exit Function
End If
Dim raw As Long, done As Long, continue_until As Long
raw = 0: done = 0: continue_until = CLng(Sub_Bank) * 256 * 48 '11600 '11776

Do
    If Asc(Mid(ROM_Data, Start_from + raw, 1)) >= 128 Then
        done = done + Asc(Mid(ROM_Data, Start_from + raw + 1, 1)) + 1
        raw = raw + 2
    Else
        raw = raw + 1
        done = done + 1
    End If
Loop Until done >= continue_until
Skip_To = Start_from + raw
End Function

Private Sub Display_block_Map(Level As Integer)
Dim Display_from As Long, offset As Long
If Store(Level).Unzipped_Modified = vbNullString Then
    Display_from = Skip_To(Store(Level).Map_Bank, Store(Level).Map_Sub_Bank)
    Parse_Map Display_from, Level
End If
Status "Displaying: Tile Map " & Level, 1

For X = 0 To 255
    For Y = 0 To 47
        Position = Y * 256 + X + 1
        tileasc = Asc(Mid(Store(Level).Unzipped_Modified, Position, 1))

        sy = Int(tileasc / 16)
        sx = tileasc - sy * 16
        
        'If Check1(2) = 1 And Map_Zoom <> 16 Then Mode = HALFTONE Else Mode = 3
        'SetStretchBltMode TileMap1.hDC, Mode
        'StretchBlt TileMap1.hDC, X * Map_Zoom, Y * Map_Zoom, Map_Zoom, Map_Zoom, Blocks1.hDC, sx * 16, sy * 16, 16, 16, SRCCOPY
        BitBlt LayerDC(1), X * 16, Y * 16, 16, 16, Blocks1.hdc, sx * 16, sy * 16, SRCCOPY
    Next Y
Next X
'TileMap1.Refresh
'DoEvents
End Sub

Private Sub Redraw_block_Map(x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer)

If Store(Level).Unzipped_Modified = vbNullString Then
    Display_block_Map CL
    Exit Sub
End If

For X = x1 To x2
    For Y = y1 To y2
        Position = Y * 256 + X + 1
        tileasc = Asc(Mid(Store(Level).Unzipped_Modified, Position, 1))

        sy = Int(tileasc / 16)
        sx = tileasc - sy * 16
        
        'If Check1(2) = 1 And Map_Zoom <> 16 Then mode = HALFTONE Else mode = 3
        'SetStretchBltMode TileMap1.hDC, 3
        'StretchBlt TileMap1.hDC, X * Map_Zoom, Y * Map_Zoom, Map_Zoom, Map_Zoom, Blocks1.hDC, sx * 16, sy * 16, 16, 16, SRCCOPY
        BitBlt LayerDC(0), X * 16, Y * 16, 16, 16, Blocks1.hdc, sx * 16, sy * 16, SRCCOPY
    Next Y
Next X
'Display_borders
'TileMap1.Refresh
'DoEvents
End Sub

Public Sub Parse_Map(parse_from As Long, Level As Integer)
Dim Parse_pos As Integer
Do
    asc_val = Asc(Mid(ROM_Data, parse_from + Parse_pos, 1))
    If asc_val >= 128 Then
        rep_chr = Chr(asc_val - 128)
        repeats = Asc(Mid(ROM_Data, parse_from + Parse_pos + 1, 1))
        Store(Level).Unzipped_Modified = Store(Level).Unzipped_Modified & String(repeats + 1, rep_chr)
        Parse_pos = Parse_pos + 2
    Else
        Store(Level).Unzipped_Modified = Store(Level).Unzipped_Modified & Mid(ROM_Data, parse_from + Parse_pos, 1)
        Parse_pos = Parse_pos + 1
    End If
Loop Until Len(Store(Level).Unzipped_Modified) >= 12288 '12800
Store(Level).Zipped = Mid(ROM_Data, parse_from, Parse_pos)
Store(Level).Unzipped = Store(Level).Unzipped_Modified
End Sub

Public Sub set_palette(c0 As Long, c1 As Long, c2 As Long, c3 As Long)
Current_Palette(0) = c0
Current_Palette(1) = c1
Current_Palette(2) = c2
Current_Palette(3) = c3
End Sub

Private Sub Load_Tiles(Level As Integer, map_tiles_start As Long, set_tiles As Long)
Dim X As Integer, Y As Integer, Tile As Integer
Status "Displaying: Tiles", 1

If map_tiles_start = -1 Or Check2(3) = 0 Then
    Random_Tiles
    Exit Sub
End If

For Y = 0 To 15
    For X = 0 To 15
        If ROM_Name = "Mario 2" Then
            If Tile > 159 Then 'map tiles
                General_Load_Tile Tiles1, map_tiles_start, Tile - 160, X, Y, True
            ElseIf Tile > 103 Then 'set tiles
                Select Case Get_Header(CL, 16)
                Case 228 ' norm
                    bricks = 0
                Case 147 'space
                    bricks = 3
                Case 225 'grey
                    bricks = 1
                Case Else
                    Status "weird palette header: '" & Get_Header(CL, 16) & "'", 4
                    bricks = 2
                End Select
                General_Load_Tile Tiles1, set_tiles, Tile - 104 + (bricks * 56), X, Y, True
            ElseIf Tile > 47 Then
                'General_Load_Tile Tiles1, 104449 + (level * 16 * 48), Tile - 48, x, y, True
            Else
                General_Load_Tile Tiles1, 454657, Tile, X, Y, True
            End If
        ElseIf ROM_Name = "Wario" Then
            If Tile > 159 Then 'map tiles
                If CL = 0 Then
                    General_Load_Tile Tiles1, map_tiles_start, Tile - 160, X, Y, True
                End If
            ElseIf Tile > 127 Then 'set tiles
                General_Load_Tile Tiles1, set_tiles, Tile - 127, X, Y, True
            ElseIf Tile > 47 Then
                'general_Load_Tile tiles1,104449 + (level * 16 * 48), Tile - 48, X, Y,true
            Else
                'general_Load_Tile tiles1,454657, Tile, x, Y,true
            End If
        End If
        Tile = Tile + 1
    Next X
Next Y

Tiles1.Refresh
End Sub

Private Function Get_Tile_Start(Level As Integer) As Long
Dim ret As String
ret = From_ini("tiles", "level" & Level & "map")
If ret <> vbNullString Then
    Get_Tile_Start = Val(ret)
    Exit Function
End If
rule = 0: multiplier = Level
Do
    rule = rule + 1
    current_rule = From_ini("tiles", rule & "end", 0, True)
    If Level <= current_rule Then
        ret = From_ini("tiles", rule & "map", -1, True)
        Get_Tile_Start = Val(ret) + (multiplier * 16 * 96)
        Exit Function
    ElseIf current_rule = 0 Then
        Get_Tile_Start = -1
        Exit Function
    Else
        multiplier = Level - current_rule - 1
    End If
Loop
End Function

Private Sub Random_Tiles()
Tiles1.Cls
For Y = 0 To 15
    For X = 0 To 15
        tileval = Y * 16 + X
        colour = RGB((tileval Mod 8) * 32, ((tileval \ 8) Mod 8) * 32, ((tileval \ 32) Mod 8) * 32)
        If tileval = 127 Then colour = RGB(255, 255, 255)
        Tiles1.Line (X * 8, Y * 8)-((X + 1) * 8, (Y + 1) * 8), colour, BF
    Next X
Next Y
Tiles1.Refresh
End Sub

Private Sub Random_Blocks()
For Y = 0 To 7
    For X = 0 To 15
        tileval = Y * 16 + X
        colour = RGB((tileval Mod 8) * 32, ((tileval \ 8) Mod 8) * 32, ((tileval \ 16) Mod 8) * 32)
        If tileval = Asc("`") Then colour = RGB(255, 255, 255)

        Blocks1.Line (X * 16, Y * 16)-((X + 1) * 16, (Y + 1) * 16), colour, BF
    Next X
Next Y
BitBlt Blocks2.hdc, 0, 0, Blocks1.ScaleWidth, Blocks1.ScaleHeight, Blocks1.hdc, 0, 0, SRCCOPY
Blocks1.Refresh: Blocks2.Refresh
End Sub

Private Sub Load_Blocks(Block_Bank As Long, blocks_start As Long)
Dim dx As Integer, dy As Integer, sx As Integer, sy As Integer

Status "Displaying: Blocks", 1
If blockbank = -1 Or Check2(2) = 0 Then
    Random_Blocks
    Exit Sub
ElseIf Check1(3) = 1 Then
    Blocks1 = LoadPicture(App.Path & "\config\" & From_ini("config", "blocks"))
    Exit Sub
End If
If Check1(1) = 1 Then
    special_blocks1 = LoadPicture(App.Path & "\config\" & From_ini("config", "blocks"))
    DoEvents
End If

If Len(Store(CL).Block_data) = 0 Then
    Store(CL).Block_data = Mid(ROM_Data, blocks_start + 256 * Block_Bank + 1, 512)
End If

For B = 0 To 127
    plain_temp = 0
    For X = 0 To 1
        For Y = 0 To 1
        tileasc = Asc(Mid(Store(CL).Block_data, B * 4 + Y * 2 + X + 1, 1)) - 128
        If tileasc < 0 Then tileasc = tileasc + 256
        sy = Int(tileasc / 16)
        sx = tileasc - sy * 16
        
        dy = Int(B / 16)
        dx = B - dy * 16
        If Plain_Tile(tileasc) Then
            plain_temp = plain_temp + 1
        End If
        BitBlt Blocks1.hdc, dx * 16 + X * 8, dy * 16 + Y * 8, 8, 8, Tiles1.hdc, sx * 8, sy * 8, SRCCOPY
        Next Y
    Next X
    If Check1(1).value = 1 Then
        flag = special_blocks1.Point(256 + dx, dy * 16 + 4)
        If flag = vbBlue Then
            Repalette dx, dy
        ElseIf flag = vbRed Or plain_temp = 4 Then
            TransparentBlt Blocks1.hdc, 16 * dx, 16 * dy, 16, 16, special_blocks1.hdc, 16 * dx, 16 * dy, 16, 16, vbGreen
        ElseIf flag = vbGreen Then
            'do nothing
        Else
            Repalette dx, dy
        End If
    End If
Next B

'If Check1(1).Value = 1 Then BitBlt Blocks1.hdc, 16 * 14, 16 * 7, 16, 16, Blocks1.hdc, 16, 0, SRCINVERT

BitBlt Blocks2.hdc, 0, 0, Blocks1.ScaleWidth, Blocks1.ScaleHeight, Blocks1.hdc, 0, 0, SRCCOPY
Blocks1.Refresh: Blocks2.Refresh
'DoEvents
End Sub

Private Sub Repalette(dx As Integer, dy As Integer)
Dim OldC As Long, NewC As Long
For X = 0 To 15
    For Y = 0 To 15
        OldC = Blocks1.Point(dx * 16 + X, dy * 16 + Y)
        cn = -1
        Do
            cn = cn + 1
        Loop Until OldC = palette1(3 - cn).BackColor Or cn = 3
        NewC = special_blocks1.Point(256 + dx, dy * 16 + cn)
        If NewC <> vbGreen Then Blocks1.PSet (dx * 16 + X, dy * 16 + Y), NewC
    Next Y
Next X
End Sub
Private Sub Load_Zone_List()
totalzones = From_ini("config", "zones", 0, True)
If totalzones > 0 Then
    For z = 1 To totalzones
        Combo1.AddItem From_ini("zone" & z, "name", "zone" & z), z - 1
    Next z
    Combo1.Enabled = True
End If
Combo1.ListIndex = 0
End Sub

Private Sub Load_Config()
If Check2(4).value = 0 Then Exit Sub
If From_ini("config", "headers", -1, True) = -1 Then
    Check2(1).value = 0
Else
    Check2(1).value = 1
End If
If From_ini("config", "blocks") = "no" Then
    Check2(2).value = 0
Else
    Check2(2).value = 1
End If
If From_ini("config", "tiles") = "no" Then
    Check2(3).value = 0
Else
    Check2(3).value = 1
End If
End Sub

Public Sub Compress_map(Level As Integer)
On Error Resume Next
Dim pos As Integer
If Store(Level).Unzipped_Modified = Store(Level).Unzipped Then
    Exit Sub
End If
Store(Level).Zipped = vbNullString
Do
    pos = pos + 1
    asc_val = Asc(Mid(Store(Level).Unzipped_Modified, pos, 1))
    repeats = 0
    If asc_val = Asc(Mid(Store(Level).Unzipped_Modified, pos + 1, 1)) Then
        Do
            repeats = repeats + 1
            test_asc = Asc(Mid(Store(Level).Unzipped_Modified, pos + repeats, 1))
            eol = 256 - ((pos - 1) Mod 256)
            
        Loop Until test_asc <> asc_val Or repeats >= eol Or pos + repeats >= Len(Store(Level).Unzipped_Modified)
    End If
    If repeats >= 3 Then
        pos = pos + repeats - 1
        Store(Level).Zipped = Store(Level).Zipped & Chr(asc_val + 128) & Chr(repeats - 1)
    Else
        Store(Level).Zipped = Store(Level).Zipped & Mid(Store(Level).Unzipped_Modified, pos, 1)
    End If
Loop Until pos = Len(Store(Level).Unzipped_Modified)
Status "New Map Size: (compressed) " & Len(Store(Level).Zipped)
End Sub
Public Sub Load_All_Levels()
Dim Level As Integer, Display_from As Long
Do
    If Store(Level).Unzipped_Modified = vbNullString Then
        Display_from = Skip_To(Store(Level).Map_Bank, Store(Level).Map_Sub_Bank)
        Parse_Map Display_from, Level
    End If
    Level = Level + 1
Loop Until Get_Header(Level, 12) > 0 Or Get_Header(Level, 11) <> Level
End Sub
Private Sub hScroll1_Scroll()
HScroll1_Change
End Sub
Private Sub VScroll1_Scroll()
VScroll1_Change
End Sub

Private Sub VScroll2_Change()
sprites2.Top = -VScroll2.value - 1
End Sub

Private Sub VScroll2_Scroll()
VScroll2_Change
End Sub

Private Sub VScroll3_Change()
sprites3.Top = -VScroll3.value
End Sub

Private Sub VScroll3_Scroll()
sprites3.Top = -VScroll3.value
End Sub
Private Sub Display_Tile_Map(map_num As Integer)
If map_num < 1 Then Exit Sub
tilemap2.Cls
Tilemap_Mode = LCase(From_ini("overworld", map_num & "mode", "0"))
If Tilemap_Mode = "version" Then
    v = "-" & Rom_Version
    Tilemap_Mode = LCase(From_ini("overworld", map_num & v & "mode", "0"))
End If

tilemap_start = From_ini("overworld", map_num & v & "map", 0, True)

Display_Tiles From_ini("overworld", map_num & v & "tiles", 0, True) + 1
ReDim TileMap_Store(0)
ReDim TileMap_Store(31, 31)

Select Case Tilemap_Mode
Case "4"
    ReDim TileMap_Store(63, 63)
    tilemap2.width = 513
    tilemap2.height = 513
    'DoEvents
    Load_Tile_Map tilemap_start, "0"
    Load_Tile_Map tilemap_start, "1"
    Load_Tile_Map tilemap_start, "2"
    Load_Tile_Map tilemap_start, "3"
Case "full"
    tilemap2.width = 257
    tilemap2.height = 257
    Load_Tile_Map tilemap_start, Tilemap_Mode
Case "line"
    tilemap2.width = 161
    tilemap2.height = 9
    Load_Tile_Map tilemap_start, Tilemap_Mode
Case "title"
    tilemap2.width = 161
    tilemap2.height = 145
    Load_Tile_Map tilemap_start, Tilemap_Mode
Case Else 'screen
    Tilemap_Mode = "screen"
    tilemap2.width = 161
    tilemap2.height = 145
    Load_Tile_Map tilemap_start, Tilemap_Mode
End Select

If Check9(2) Then
    Dim s As Integer
    Paths(1).X = 17: Paths(1).Y = 41
    Paths(32).X = 9: Paths(32).Y = 4
    Paths(41).X = 12: Paths(41).Y = 8
    Paths(54).X = 11: Paths(54).Y = 7
    s = From_ini("overworld", map_num & "path", "0", True)
    If s > 0 Then Load_Paths s
End If
tilemap2.Refresh
DoEvents
End Sub
Private Sub Load_Paths(c As Integer)
Dim D As Integer
Paths(c).Loading = True
'Do
    For D = 0 To 3 'direction
        Load_Path c, D
        linked = Paths(c).Compass(D).To
        If linked > 0 Then
            If Paths(linked).Loading = False Then Load_Paths (linked)
        End If
    Next D
'Loop Until c <= 0
End Sub
Private Function Calculate_Path_Location(c As Integer, D As Integer) As Long
Dim Pointer As String * 2, B As Byte, A As Byte, BA As Byte, BB As Byte
    
Pointer = Mid$(ROM_Data, 398907 + c * 8 + D * 2, 2)
If Pointer = Chr(255) & Chr(255) Then
    Calculate_Path_Location = 0
Else
    A = Asc(Mid$(Pointer, 1, 1))
    B = Asc(Mid$(Pointer, 2, 1))
    BA = Int(B / 16)
    BB = B - (BA * 16)
    Calculate_Path_Location = 376832 + (4096 * BA) + A + BB * 256
End If
End Function
Private Sub Load_Path(Cross As Integer, Direction As Integer)
Dim L As Boolean, R As Boolean, U As Boolean, D As Boolean
ColourBuffer PathInfoDC, 128 * 8, 128 * 8, vbGreen

Path = Calculate_Path_Location(Cross, Direction)
If Path = 0 Then Exit Sub
X = Paths(Cross).X: Y = Paths(Cross).Y
Do
    D = False: U = False: L = False: R = False
    Path = Path + 1
    step = Asc(Mid$(ROM_Data, Path, 1))
    M = step Mod 16
    s = (step - M) / 16
    If M >= 8 Then 'down
        M = M - 8
        D = True
    End If
    If M >= 4 Then  'up
        M = M - 4
        U = True
    End If
    If M >= 2 Then 'left
        M = M - 2
        L = True
    End If
    If M >= 1 Then R = True

    Select Case s
    Case 0, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14
        speed = 1
    Case 1, 10, 15
        speed = 0.5
    End Select
    
    If R Then
        BitBlt PathInfoDC, X * 8, Y * 8, 8, 8, special_blocks1.hdc, 36 * 8, 0, SRCCOPY
        X = X + speed
    ElseIf L Then
        BitBlt PathInfoDC, X * 8, Y * 8, 8, 8, special_blocks1.hdc, 37 * 8, 0, SRCCOPY
        X = X - speed
    End If
    If U Then
        BitBlt PathInfoDC, X * 8, Y * 8, 8, 8, special_blocks1.hdc, 38 * 8, 0, SRCCOPY
        Y = Y - speed
    ElseIf D Then
        BitBlt PathInfoDC, X * 8, Y * 8, 8, 8, special_blocks1.hdc, 39 * 8, 0, SRCCOPY
        Y = Y + speed
    End If
Loop Until step = 0

c = Asc(Mid(ROM_Data, Path + 1, 1))
If Paths(c).Loading = False Then
    Paths(c).X = X: Paths(c).Y = Y
End If
Paths(Cross).Compass(Direction).To = c
'BitBlt PathInfoDC, X * 8, Y * 8, 8, 8, special_blocks1.hdc, 36 * 8, 0, SRCCOPY

TransBltOverlay tilemap2.hdc, PathInfoDC, 0, 0, 128 * 8, 128 * 8, vbGreen
End Sub


Private Sub Load_Tile_Map(Map_Start As Long, Optional drawmode As String)
For Y = 0 To 31
    For X = 0 To 31
        Select Case drawmode
        Case "title"
            If Y < 18 And X < 28 Then
                Position = Y * 32 + X + 8
                tileasc = Asc(Mid(ROM_Data, Map_Start + Position, 1))
                TileMap_Store(X, Y) = tileasc
            End If
        Case "full", "0", "1", "2", "3"
            box = Val(drawmode)
            box_x = box Mod 2
            box_y = (box - box_x) \ 2
            
            Position = (box * 32 * 32) + Y * 32 + X + 1
            tileasc = Asc(Mid(ROM_Data, Map_Start + Position, 1))
            TileMap_Store(box_x * 32 + X, box_y * 32 + Y) = tileasc
        Case "line"
            If Y = 0 And X < 20 Then
                Position = Y * 32 + X + 1
                tileasc = Asc(Mid(ROM_Data, Map_Start + Position, 1))
                TileMap_Store(X, Y) = tileasc
            End If
        Case Else 'screen
            If Y < 18 And X < 20 Then
                Position = Y * 32 + X + 1
                tileasc = Asc(Mid(ROM_Data, Map_Start + Position, 1))
                TileMap_Store(X, Y) = tileasc
            End If
        End Select
    Next X
Next Y
draw_Tile_Map drawmode
End Sub

Private Sub draw_Tile_Map(Optional drawmode As String)
ubound_x = tilemap2.ScaleWidth \ 8
ubound_y = tilemap2.ScaleHeight \ 8


For Y = 0 To ubound_y
    For X = 0 To ubound_x
        Position = Y * 32 + X + 1
        tileasc = TileMap_Store(X, Y)

        sy = Int(tileasc / 16)
        sx = tileasc - sy * 16
        
        sy = sy - 8
        If sy < 0 Then sy = sy + 16
        BitBlt tilemap2.hdc, X * 8, Y * 8, 8, 8, tiles2.hdc, sx * 8, sy * 8, SRCCOPY
    Next X
Next Y
'tilemap2.Refresh
'DoEvents
End Sub
Private Sub Display_Tiles(tiles_start As Long)
Dim X As Integer, Y As Integer
For X = 0 To 15
    For Y = 0 To 15
        General_Load_Tile tiles2, tiles_start, Y * 16 + X, X, Y
    Next Y
Next X
tiles2.Refresh
End Sub

Public Function Export_Header(Export_Type As String, Export_Size As Long, Optional Export_level As Integer) As String
Export_Header = "<"

Export_Header = Export_Header & "Type=" & Export_Type
Export_Header = Export_Header & " Size=" & Export_Size
Export_Header = Export_Header & " Verion=" & App.Major & App.Minor
If Export_level > 0 Then Export_Header = Export_Header & " Level=" & Export_level

Export_Header = Export_Header & ">"
End Function
Private Function Parse_Credit_Text(to_parse As String) As String
Dim Finish As Integer, Start As Integer, temp As String, Line_Temp As String
Do
    Start = Finish + 1
    Finish = InStr(Start, to_parse, vbCrLf)
    If Finish = 0 Then Exit Do
    Line_Temp = Mid(to_parse, Start + 1, Finish - Start - 1)
    If Mid(Line_Temp, 1, 6) = "<FADE>" Then
        'nop
    ElseIf Len(Line_Temp) > 20 Then
        temp = temp & Left(Line_Temp, 20)
    ElseIf Len(Line_Temp) < 20 Then
        temp = temp & Line_Temp & String(20 - Len(Line_Temp), " ")
    Else '=20
        temp = temp & Line_Temp
    End If
Loop Until Finish >= Len(to_parse)
Parse_Credit_Text = temp
End Function

