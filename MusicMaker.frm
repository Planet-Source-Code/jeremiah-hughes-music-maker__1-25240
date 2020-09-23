VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Music Maker"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   -1335
   ClientWidth     =   10650
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "MusicMaker.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   565
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   710
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton RewindBut 
      Height          =   375
      Left            =   360
      Picture         =   "MusicMaker.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "Return to Start"
      Top             =   840
      Width           =   375
   End
   Begin VB.Frame FunctionsFrame 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   120
      TabIndex        =   37
      Top             =   240
      Width           =   3015
      Begin VB.CommandButton InsertColumnBut 
         Height          =   375
         Left            =   2520
         Picture         =   "MusicMaker.frx":0BF2
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Insert Space"
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton PasteIntoBut 
         Height          =   375
         Left            =   2160
         Picture         =   "MusicMaker.frx":0D44
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Paste (Insert)"
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton PasteBut 
         Height          =   375
         Left            =   1800
         Picture         =   "MusicMaker.frx":14F4
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Paste (Overwrite)"
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton CutBut 
         Height          =   375
         Left            =   1440
         Picture         =   "MusicMaker.frx":1CA4
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Cut Selection"
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton CopyBut 
         Height          =   375
         Left            =   1080
         Picture         =   "MusicMaker.frx":2454
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Copy Selection"
         Top             =   0
         Width           =   375
      End
      Begin VB.OptionButton FunctionBut 
         Height          =   375
         Index           =   2
         Left            =   720
         Picture         =   "MusicMaker.frx":2C04
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Select Area"
         Top             =   0
         Width           =   375
      End
      Begin VB.OptionButton FunctionBut 
         Height          =   375
         Index           =   1
         Left            =   360
         Picture         =   "MusicMaker.frx":33B2
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Start/Insertion Position"
         Top             =   0
         Width           =   375
      End
      Begin VB.OptionButton FunctionBut 
         Height          =   375
         Index           =   0
         Left            =   0
         Picture         =   "MusicMaker.frx":3B60
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Draw Mode"
         Top             =   0
         Value           =   -1  'True
         Width           =   375
      End
   End
   Begin VB.Frame TransposeFrame 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   9120
      TabIndex        =   34
      Top             =   240
      Width           =   1215
      Begin VB.CommandButton TransposeBut 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   600
         Picture         =   "MusicMaker.frx":4310
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Move All Notes Up One"
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton TransposeBut 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         Picture         =   "MusicMaker.frx":4AC0
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Move All Notes Down One"
         Top             =   0
         Width           =   375
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9720
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.mia"
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3120
      TabIndex        =   31
      Top             =   1080
      Width           =   2895
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      LargeChange     =   10
      Left            =   3120
      Max             =   300
      Min             =   20
      TabIndex        =   28
      Top             =   240
      Value           =   120
      Width           =   2895
   End
   Begin VB.OptionButton PlayAndStop 
      Height          =   375
      Index           =   1
      Left            =   1320
      Picture         =   "MusicMaker.frx":5270
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Stop"
      Top             =   840
      Value           =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton PlayAndStop 
      Height          =   375
      Index           =   0
      Left            =   840
      Picture         =   "MusicMaker.frx":5A20
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Play"
      Top             =   840
      Width           =   375
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   6015
      LargeChange     =   25
      Left            =   10260
      Max             =   115
      Min             =   12
      TabIndex        =   24
      Top             =   2040
      Value           =   12
      Width           =   255
   End
   Begin VB.Frame MuteTrackFrame 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   15
      Top             =   1080
      Width           =   2895
      Begin VB.CheckBox TrackMute 
         BackColor       =   &H00446699&
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Serpentine"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   0
         Width           =   375
      End
      Begin VB.CheckBox TrackMute 
         BackColor       =   &H00FF00FF&
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Serpentine"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   0
         Width           =   375
      End
      Begin VB.CheckBox TrackMute 
         BackColor       =   &H00FF8080&
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Serpentine"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   0
         Width           =   375
      End
      Begin VB.CheckBox TrackMute 
         BackColor       =   &H00FFFF00&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Serpentine"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   0
         Width           =   375
      End
      Begin VB.CheckBox TrackMute 
         BackColor       =   &H0000FF00&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Serpentine"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   0
         Width           =   375
      End
      Begin VB.CheckBox TrackMute 
         BackColor       =   &H0000C0C0&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Serpentine"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   0
         Width           =   375
      End
      Begin VB.CheckBox TrackMute 
         BackColor       =   &H000080FF&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Serpentine"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   0
         Width           =   375
      End
      Begin VB.CheckBox TrackMute 
         BackColor       =   &H000000FF&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Serpentine"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox MusicBox 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H000000FF&
      FillStyle       =   7  'Diagonal Cross
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   600
      ScaleHeight     =   401
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   641
      TabIndex        =   0
      Top             =   2040
      Width           =   9615
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   40
      Left            =   600
      Max             =   980
      Min             =   20
      TabIndex        =   1
      Top             =   8100
      Value           =   20
      Width           =   9645
   End
   Begin VB.PictureBox SwapScreen 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   600
      ScaleHeight     =   401
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   641
      TabIndex        =   3
      Top             =   2040
      Visible         =   0   'False
      Width           =   9615
   End
   Begin VB.Frame EditTrackFrame 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   4
      Top             =   240
      Width           =   2895
      Begin VB.OptionButton TrackSel 
         BackColor       =   &H00446699&
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Serpentine"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         Width           =   375
      End
      Begin VB.OptionButton TrackSel 
         BackColor       =   &H00FF00FF&
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Serpentine"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   0
         Width           =   375
      End
      Begin VB.OptionButton TrackSel 
         BackColor       =   &H00FF8080&
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Serpentine"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         Width           =   375
      End
      Begin VB.OptionButton TrackSel 
         BackColor       =   &H00FFFF00&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Serpentine"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         Width           =   375
      End
      Begin VB.OptionButton TrackSel 
         BackColor       =   &H0000FF00&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Serpentine"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         Width           =   375
      End
      Begin VB.OptionButton TrackSel 
         BackColor       =   &H0000C0C0&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Serpentine"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   375
      End
      Begin VB.OptionButton TrackSel 
         BackColor       =   &H000080FF&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Serpentine"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   375
      End
      Begin VB.OptionButton TrackSel 
         BackColor       =   &H000000FF&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Serpentine"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Value           =   -1  'True
         Width           =   375
      End
   End
   Begin VB.Label TopBarOverlay 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   58
      Top             =   1800
      Width           =   9600
   End
   Begin VB.Label TopBar 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   18
      Left            =   9720
      TabIndex        =   67
      Top             =   1800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label TopBar 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   17
      Left            =   9720
      TabIndex        =   66
      Top             =   1800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label TopBar 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   16
      Left            =   9720
      TabIndex        =   65
      Top             =   1800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label TopBar 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   15
      Left            =   9720
      TabIndex        =   64
      Top             =   1800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label TopBar 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   14
      Left            =   9720
      TabIndex        =   63
      Top             =   1800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label TopBar 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   13
      Left            =   9720
      TabIndex        =   62
      Top             =   1800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label TopBar 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   12
      Left            =   9720
      TabIndex        =   61
      Top             =   1800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label TopBar 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   11
      Left            =   9720
      TabIndex        =   60
      Top             =   1800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label TopBar 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   9720
      TabIndex        =   59
      Top             =   1800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label TopBar 
      BackColor       =   &H00000000&
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   57
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label TopBar 
      BackColor       =   &H00000000&
      Caption         =   "1"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   56
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label TopBar 
      BackColor       =   &H00000000&
      Caption         =   "2"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   55
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label TopBar 
      BackColor       =   &H00000000&
      Caption         =   "3"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   3480
      TabIndex        =   54
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label TopBar 
      BackColor       =   &H00000000&
      Caption         =   "4"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   4440
      TabIndex        =   53
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label TopBar 
      BackColor       =   &H00000000&
      Caption         =   "5"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   5400
      TabIndex        =   52
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label TopBar 
      BackColor       =   &H00000000&
      Caption         =   "6"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   6360
      TabIndex        =   51
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label TopBar 
      BackColor       =   &H00000000&
      Caption         =   "7"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   7320
      TabIndex        =   50
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label TopBar 
      BackColor       =   &H00000000&
      Caption         =   "8"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   8280
      TabIndex        =   49
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label TopBar 
      BackColor       =   &H00000000&
      Caption         =   "9"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   9240
      TabIndex        =   48
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Functions"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   0
      Width           =   2895
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Transpose"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9120
      TabIndex        =   33
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label InstLabel 
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Instruments"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3120
      TabIndex        =   30
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label TempoLabel 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   3120
      TabIndex        =   29
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Tempo"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3120
      TabIndex        =   27
      Top             =   0
      Width           =   2895
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   120
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Mute Tracks"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6120
      TabIndex        =   14
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Edit Track"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6120
      TabIndex        =   13
      Top             =   0
      Width           =   2895
   End
   Begin VB.Shape Shape1 
      Height          =   6045
      Left            =   120
      Top             =   2040
      Width           =   10125
   End
   Begin VB.Label SideBar 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   6315
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   600
      Top             =   1800
      Width           =   9915
   End
   Begin VB.Menu mFile 
      Caption         =   "File"
      Begin VB.Menu mNew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mOpen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mSaveAs 
         Caption         =   "Save as..."
      End
   End
   Begin VB.Menu midi_devices 
      Caption         =   "Midi Device"
      Begin VB.Menu Device 
         Caption         =   ""
         Index           =   0
      End
      Begin VB.Menu Device 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu Device 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu Device 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu Device 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu Device 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu Device 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu Device 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu Device 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu Device 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   9
         Visible         =   0   'False
      End
      Begin VB.Menu Device 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   10
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Dim Grid%(7, 1000) '(Tracks, Value)
Dim InstGrid%(7, 1000)
Dim CopyGrid%(7, 1000)
Dim CopyInstGrid%(7, 1000)
Dim CopyGridLen%
Dim A%, B%, C%, D%, E%, F%, G%, Z%
Dim XSize%, YSize%
Dim XSizeMid%, YSizeMid%
Dim Temp$
Dim NoteNames$(127), Sharps(100) As Boolean
Dim Colr As Long, BGColr As Long
Dim GridX%, GridY%
Dim OldGridX%, OldGridY%
Dim MovingGridX%, OldMovingGridX%
Dim HoverX%, HoverY%
Dim OldHoverX%, OldHoverY%
Dim StartX%, StartY%
Dim EndX%, EndY%
Dim NotePlayX%
Dim Note%
Dim CurrentTrack%, TrackColr As Long
Dim TrackColrs(7) As Long, DimmedColrs(7) As Long
Dim TheNotes As Variant
Dim Offset%
Dim Flip%
Dim TopBarOffset%
Dim Tick As Long
Dim PlayingSong As Boolean, PlayX%
Dim OldNote%(7)
Dim NoteType%, OldNoteType%
Dim TempoWait%
Dim CurrentInst%(7), OldInst%(7)
Dim SongLength%, TrackLength%
Dim FilePath$
Dim FF%
Dim CurrentlyOpenFile
Dim CursorType As Integer, ColumnX As Integer
Dim PlayBarX As Integer, HasStarted As Boolean
Dim JustScrolled As Boolean, TurningPage As Boolean
Dim MouseIsDown As Boolean
Dim SelectStartX%, SelectEndX%
Dim VisSelStartX%, VisSelEndX%
Dim TempColumnCursor As Boolean
Dim Resizing As Boolean

Const DrawCursor = 0
Const ColumnCursor = 1
Const SelectCursor = 2

Const BlankNote = 0
Const StartingNote = 1
Const ContinuingNote = 2

'Midi Variables (from midi piano program)

'for piano play
Dim numDevices As Long      ' number of midi output devices
Dim curDevice As Long       ' current midi device
Dim hmidi As Long           ' midi output handle
Dim rc As Long              ' return code
Dim midimsg As Long         ' midi output message buffer
Dim channel As Integer      ' midi output channel
Dim volume As Integer       ' midi volume
Dim baseNote As Integer     ' the first note on our "piano"
Dim incra As Integer        ' increment the note
Dim Tempo As Integer        ' set playing speed
Dim incraup As Integer      ' incra-1
Dim curdr As String         ' gets current directory

Private Sub Combo1_Click()
Dim Instrument%

Instrument = Combo1.ListIndex

'if anything is selected then change all selected notes on current track
'to selected instrument, otherwise just change current instrument
If SelectStartX > -1 Then
    For A = SelectStartX To SelectEndX
        If Grid(CurrentTrack, A) > -1 Then
            InstGrid(CurrentTrack, A) = Instrument
        End If
    Next A
End If

CurrentInst(CurrentTrack) = Instrument
channel = CurrentTrack
ChangeInstrument CurrentInst(CurrentTrack)

End Sub

Private Sub CopyBut_Click()
MusicBox.SetFocus

If SelectStartX > -1 Then 'something is selected
    
    CopySelectedArea
    
    SelectStartX = -1 'remove blue highlight
    SelectEndX = -1

    DrawGrid
    
    FunctionBut(1).Value = True 'change to red column mode
    
End If

End Sub

Private Sub CutBut_Click()
Dim Jump%

MusicBox.SetFocus

If SelectStartX > -1 Then 'something is selected
    
    CopySelectedArea
        
    'move all notes that are after selected area to start of selected area
    Jump = CopyGridLen + 1
    For A = SelectStartX To 1000 - Jump
        For B = 0 To 7
            Grid(B, A) = Grid(B, A + Jump)
            InstGrid(B, A) = InstGrid(B, A + Jump)
        Next B
    Next A
    
    'change start of severed notes from continuing notes to starting notes
    For A = 0 To 7
        If Grid(A, SelectStartX) >= 1000 Then
            Grid(A, SelectStartX) = Grid(A, SelectStartX) - 1000
        End If
    Next A

    'delete extra space at end of Grid and InstGrid
    For A = 1001 - Jump To 1000
        For B = 0 To 7
            Grid(B, A) = -1
            InstGrid(B, A) = 0
        Next B
    Next A
    
    SelectStartX = -1 'remove blue highlight
    SelectEndX = -1
    
    FunctionBut(1).Value = True 'change to red column mode

    DrawGrid
    
End If

End Sub

Private Sub CopySelectedArea()

CopyGridLen = SelectEndX - SelectStartX 'record length of selected area
Erase CopyGrid, CopyInstGrid

'copy selected area's data to CopyGrid and CopyInstGrid
For A = SelectStartX To SelectEndX
    For B = 0 To 7
        CopyGrid(B, A - SelectStartX) = Grid(B, A)
        CopyInstGrid(B, A - SelectStartX) = InstGrid(B, A)
    Next B
Next A

'change start of severed notes from continuing notes to starting notes
For A = 0 To 7
    If CopyGrid(A, 0) >= 1000 Then
        CopyGrid(A, 0) = CopyGrid(A, 0) - 1000
    End If
Next A

End Sub


Private Sub device_Click(Index As Integer)

Device(curDevice + 1).Checked = False
Device(Index).Checked = True
curDevice = Index - 1
rc = midiOutClose(hmidi)
rc = midiOutOpen(hmidi, curDevice, 0, 0, 0)
   
If (rc <> 0) Then
      MsgBox "Couldn't open midi out, rc = " & rc
End If

End Sub

Private Sub Form_Load()

'Midi Stuff
   Dim I As Long
   Dim caps As MIDIOUTCAPS
   
   ' Set the first device as midi mapper
   Device(0).Caption = "MIDI Mapper"
   Device(0).Visible = True
   Device(0).Enabled = True
   
   ' Get the rest of the midi devices
   numDevices = midiOutGetNumDevs()
   
   For I = 0 To (numDevices - 1)
      midiOutGetDevCaps I, caps, Len(caps)
        Device(I + 1).Caption = caps.szPname
        Device(I + 1).Visible = True
      Device(I + 1).Enabled = True
   Next
   
   'Select the MIDI Mapper as the default device
   device_Click (0)
   
   ' Set the default channel
   channel = 0
   
   ' Set volume range
   volume = 127



'Initialize
ClearGrids

For A = 100 To 227 'load up instrument names
    Combo1.AddItem A - 100 & " " & LoadResString(A)
Next A

'I haven't quite figured out drums yet...
'For A = 35 To 81 'load up drum names
'    Combo1.AddItem A + 93 & " " & LoadResString(A)
'Next A

Combo1.ListIndex = 0

XSize = 39 'width of musicbox (columns)
YSize = 24 'height of musicbox (rows)
XSizeMid = (XSize + 1) \ 2
YSizeMid = (YSize + 1) \ 2

CursorType = 0 'draw mode
ColumnX = 0 'red play start position bar starts at 0
StartX = 0 'leftmost column displayed
EndX = StartX + XSize 'rightmost column displayed
StartY = 55 'topmost row displayed
EndY = StartY + YSize 'bottommost row displayed
SelectStartX = -1 'no area is selected
SelectEndX = -1 'no area is selected

TopBarOffset = 40 'for printing topbar (column #'s) 40 pixles from left edge of form

Tempo = 120
HScroll2.Value = Tempo

CurrentTrack = 0
For A = 0 To 7
    TrackColrs(A) = TrackSel(A).BackColor
Next A
TrackColr = TrackColrs(0)

DimmedColrs(0) = &HC0C0FF
DimmedColrs(1) = &HC0E0FF
DimmedColrs(2) = &HC0FFFF
DimmedColrs(3) = &HC0FFC0
DimmedColrs(4) = &HFFFFC0
DimmedColrs(5) = &HFFC0C0
DimmedColrs(6) = &HFFC0FF
DimmedColrs(7) = &HAACCEE
BGColr = MusicBox.BackColor

'Assign names to all 128 notes
'notes go from high 10G (note 0) to low 0C (note 128)
'to better suit screen's top to bottom Y coordinate system
'when playing note, number is flipped (127 - #)
TheNotes = Split("C,C#,D,D#,E,F,F#,G,G#,A,A#,B", ",")
For A = 0 To 11
    For B = 0 To 11
        C = A * 12 + B
        If C < 128 Then
            NoteNames(127 - C) = A & TheNotes(B)
        End If
    Next B
Next A

PrintSideBar

DrawGrid

'check to see if file should be loaded on startup
If Len(Command) > 2 Then
    FilePath = Mid(Command, 2, Len(Command) - 2)
    If Dir(FilePath) <> "" Then
        CurrentlyOpenFile = FilePath
        OpenFile
    End If
End If

End Sub

Private Sub Form_Resize()

Resizing = True

If Me.WindowState = 1 Then Exit Sub


If WindowState <> 2 Then
    If Me.ScaleWidth < 718 Then
        Me.Width = ScaleX(718, vbPixels, vbTwips)
    End If

    If Me.ScaleHeight < 325 Then
        Me.Height = ScaleY(325, vbPixels, vbTwips)
    End If
End If

XSize = (Me.ScaleWidth - 70) \ 16 - 1
YSize = (Me.ScaleHeight - 165) \ 16 - 1

MusicBox.Width = (XSize + 1) * 16 + 1 'width of musicbox (pixels)
MusicBox.Height = (YSize + 1) * 16 + 1 'height of musicbox (pixels)
SwapScreen.Width = (XSize + 1) * 16 + 1
SwapScreen.Height = (YSize + 1) * 16 + 1
TopBarOverlay.Width = MusicBox.Width - 1
Shape1.Width = MusicBox.Width + 34
Shape1.Height = MusicBox.Height + 2
Shape2.Width = MusicBox.Width + 20
SideBar.Height = MusicBox.Height + 20
HScroll1.Width = MusicBox.Width + 2
HScroll1.Top = MusicBox.Height + 139
VScroll1.Height = MusicBox.Height
VScroll1.Left = MusicBox.Width + 43

XSizeMid = (XSize + 1) \ 2
YSizeMid = (YSize + 1) \ 2
'StartX = 0 'leftmost column displayed
EndX = StartX + XSize 'rightmost column displayed
StartY = 55 'topmost row displayed
EndY = StartY + YSize 'bottommost row displayed
PrintSideBar

HScroll1.Min = XSizeMid
HScroll1.Max = 999 - (XSize - XSizeMid)
HScroll1.LargeChange = XSize + 1
HScroll1.Value = XSizeMid 'value of horizontal scrollbar = middle column # of displayed screen
VScroll1.Min = YSizeMid
VScroll1.Max = 127 - YSizeMid
VScroll1.LargeChange = YSize + 1
VScroll1.Value = StartY + YSizeMid 'value of vertical scrollbar = middle row# of displayed screen

DrawGrid

Resizing = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
' Close current midi device
rc = midiOutClose(hmidi)

End Sub


Private Sub FunctionBut_Click(Index As Integer)
MusicBox.SetFocus

CursorType = Index

Select Case CursorType

Case Is = 0
    MusicBox.MousePointer = 0

Case Is = 1
    MusicBox.MousePointer = 2

Case Is = 2
    MusicBox.MousePointer = 2

End Select

End Sub

Private Sub FunctionBut_GotFocus(Index As Integer)
MusicBox.SetFocus

End Sub

Private Sub HScroll1_Change()

If Resizing Then Exit Sub

StartX = HScroll1.Value - XSizeMid
EndX = StartX + XSize
PrintTopBar
If Not TurningPage Then DrawGrid
JustScrolled = True

End Sub

Private Sub HScroll1_Scroll()
HScroll1_Change
End Sub

Private Sub HScroll2_Change()

Tempo = HScroll2.Value
TempoWait = 1000 \ (Tempo / 15)
TempoLabel.Caption = Tempo

End Sub

Private Sub HScroll2_Scroll()
HScroll2_Change
End Sub

Private Sub mNew_Click()
Dim M As VbMsgBoxResult
        
M = MsgBox("Are you sure you want to clear the board and start a new song?", vbOKCancel)

If M = vbOK Then
    CurrentlyOpenFile = ""
    ClearGrids
    DrawGrid
    Me.Caption = "Music Maker"
End If
        
End Sub

Private Sub mOpen_Click()

With CommonDialog1
    .FileName = "*.mia"
    .Filter = "*.mia"
    .DialogTitle = "Open Music File"
    .ShowOpen
    FilePath = .FileName
End With

If Right(FilePath, 5) = "*.mia" Or FilePath = "" Then 'cancel button pressed
    Exit Sub
End If

If Dir(FilePath) <> "" Then
    CurrentlyOpenFile = FilePath
    OpenFile
End If
    
End Sub

Private Sub mSave_Click()

If CurrentlyOpenFile = "" Then
    mSaveAs_Click
    Exit Sub
End If

FilePath = CurrentlyOpenFile
SaveFile

End Sub

Private Sub mSaveAs_Click()

'file DialogBox settings
With CommonDialog1
    .FileName = "*.mia"
    .Filter = "*.mia"
    .DialogTitle = "Save Music File as"
    .ShowSave
    FilePath = .FileName
End With

'cancel button pressed
If Right(FilePath, 5) = "*.mia" Or FilePath = "" Then
    Exit Sub
End If

'if file doesn't end with ".mia" then add it
If Right(FilePath, 4) <> ".mia" Then FilePath = FilePath & ".mia"

SaveFile

End Sub

Private Sub SaveFile()

FF = FreeFile

Open FilePath For Output As #FF

For A = 0 To 7
    
    TrackLength = 0
    For B = 1000 To 0 Step -1 'find length of track
        If Grid(A, B) > -1 Then
            TrackLength = B
            Exit For
        End If
    Next B
    
    Temp = ""
    For B = 0 To TrackLength
        Temp = Temp & Grid(A, B) & ","
    Next B
    
    Print #FF, Left(Temp, Len(Temp) - 1)
    
    Temp = ""
    For B = 0 To TrackLength
        Temp = Temp & InstGrid(A, B) & ","
    Next B
    
    Print #FF, Left(Temp, Len(Temp) - 1)
    
Next A

Print #FF, Tempo

Close #FF

CurrentlyOpenFile = FilePath
'MsgBox "Saved " & FilePath

Me.Caption = "Music Maker - " & FilePath

End Sub

Private Sub OpenFile()
Dim V As Variant

ClearGrids

'load file
FF = FreeFile

Open FilePath For Input As #FF

For A = 0 To 7
    Line Input #FF, Temp
    V = Split(Temp, ",")
    
    For B = 0 To UBound(V)
        Grid(A, B) = Val(V(B))
    Next B
    
    Erase V
    
    Line Input #FF, Temp
    V = Split(Temp, ",")
    
    For B = 0 To UBound(V)
        InstGrid(A, B) = Val(V(B))
    Next B
Next A

If Not EOF(FF) Then
    Input #FF, Tempo
    HScroll2.Value = Tempo
End If

Close #FF

StartX = 0
EndX = XSize
ColumnX = 0
DrawGrid

Me.Caption = "Music Maker - " & FilePath

End Sub

Private Sub MusicBox_KeyDown(KeyCode As Integer, Shift As Integer)

'if in drawmode and Ctrl key is held down
'If CursorType = DrawCursor And Shift = 2 And TempColumnCursor = False Then
'    TempColumnCursor = True
'    CursorType = ColumnCursor
'    MusicBox.MousePointer = 2
'End If

End Sub


Private Sub MusicBox_KeyUp(KeyCode As Integer, Shift As Integer)

'if in TempColumnCursor mode and Ctrl key is released
'If TempColumnCursor = True And Shift = 0 Then
'    TempColumnCursor = False
'    CursorType = DrawCursor
'    MusicBox.MousePointer = 0
'End If

End Sub

Private Sub MusicBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'prevents note from being drawn in MouseMove and MouseUp when a file was
'opened by a double-click and cursor was over MusicBox when button was
'released (MouseUp activated with no MouseDown)
MouseIsDown = True

GridX = X \ 16
GridY = Y \ 16

Select Case CursorType

'*** Draw Mode ***
Case Is = DrawCursor

    'New note
    If Button = 1 Then

        'play notes
        channel = CurrentTrack
        If TrackMute(CurrentTrack) = 0 Then StartNote StartY + GridY 'play note outside startnotes sub because note hasn't been set yet
        NotePlayX = GridX 'keep current x coordinate
        StartNotes
    
        'draw square
        GetRow (GridY)
        MusicBox.Line (GridX * 16 + 1, GridY * 16 + 1)-(GridX * 16 + 15, GridY * 16 + 15), BGColr, BF 'draw white erasing square
        MusicBox.Line (GridX * 16 + 2, GridY * 16 + 2)-(GridX * 16 + 14, GridY * 16 + 14), TrackColr, B 'draw colored hollow square
        OldGridX = GridX 'record starting point for new note

    End If

    'Erase note
    If Button = 2 Then
        If Grid(CurrentTrack, StartX + GridX) = -1 Then Exit Sub 'if trying to delete blank spot
        Do While Grid(CurrentTrack, StartX + GridX) >= 1000 'find start of note
            GridX = GridX - 1
        Loop
        Grid(CurrentTrack, GridX + StartX) = -1 'erase start of note
        InstGrid(CurrentTrack, GridX + StartX) = 0 'erase instrument
    
        GridX = GridX + 1
    
        Do While Grid(CurrentTrack, GridX + StartX) >= 1000 'find all continuations of note
            Grid(CurrentTrack, GridX + StartX) = -1 'erase continuation of note
            InstGrid(CurrentTrack, GridX + StartX) = 0 'erase instrument
            GridX = GridX + 1
        Loop
    
        DrawGrid
    End If


'*** Red Column ****
Case Is = ColumnCursor

    ColumnX = StartX + GridX
    DrawGrid


'*** Select Mode ***
Case Is = SelectCursor
    
    If Button = 1 Then
    
        If SelectStartX = -1 Or StartX + GridX < SelectStartX Then 'if no area selected or clicked behind selected area
        
            'start selecting new area
            
            'reset SelectStartX and SelectEndX
            SelectStartX = -1
            SelectEndX = -1
            
            'reset MusicBox window and get background
            DrawGrid
            GetBox
        
            'draw rectangle
            MusicBox.Line (GridX * 16, 0)-(GridX * 16 + 16, MusicBox.Height - 1), vbBlue, B
            OldGridX = GridX 'record starting point for selected area

            'record selected area start position
            SelectStartX = StartX + GridX
        
        Else
        
            'keep SelectStartX but get new SelectEndX
            
            'reset MusicBox window and get background
            DrawGrid
            GetBox
        
            'draw rectangle
            MusicBox.Line (GridX * 16, 0)-(GridX * 16 + 16, MusicBox.Height - 1), vbBlue, B
            OldGridX = GridX 'record starting point for selected area

        End If
        
        
    End If
        
    'unselect
    If Button = 2 Then
    
        'erase any previously selected area
        SelectStartX = -1
        SelectEndX = -1
        
        DrawGrid

    End If

End Select

End Sub

Private Sub MusicBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'prevents note from being drawn in MouseMove and MouseUp when a file was
'opened by a double-click and cursor was over MusicBox when button was
'released (MouseUp activated with no MouseDown)
If MouseIsDown = False And Button > 0 Then Exit Sub

Select Case CursorType

'*** Draw Mode ***
Case Is = DrawCursor

    'display instrument of current track under cursor
    If Button = 0 Then

        HoverX = X \ 16
    
        If HoverX = OldHoverX Then Exit Sub
    
        If Grid(CurrentTrack, StartX + HoverX) > -1 Then
            InstLabel.Caption = Combo1.List(InstGrid(CurrentTrack, StartX + HoverX))
        Else
            InstLabel.Caption = ""
        End If
        
        OldHoverX = HoverX
    End If
    
    'continue drawing note
    If Button = 1 Then

        MovingGridX = X \ 16

        If MovingGridX = OldMovingGridX Then Exit Sub 'if cursor hasn't moved to another square
        OldMovingGridX = MovingGridX 'remember cursor position
    
        If MovingGridX > XSize Then MovingGridX = XSize 'prevent from going past edge of screen

        If MovingGridX < GridX Then MovingGridX = GridX 'if cursor is behind starting point

        PutRow (GridY)

        MusicBox.Line (GridX * 16 + 1, GridY * 16 + 1)-(MovingGridX * 16 + 14, GridY * 16 + 14), BGColr, BF
        MusicBox.Line (GridX * 16 + 2, GridY * 16 + 2)-(MovingGridX * 16 + 14, GridY * 16 + 14), TrackColr, B
        
    End If
    
    
'*** Select Mode ***
Case Is = SelectCursor

    'continue selecting area
    If Button = 1 Then

        MovingGridX = X \ 16

        If MovingGridX = OldMovingGridX Then Exit Sub 'if cursor hasn't moved to another column
        OldMovingGridX = MovingGridX 'remember cursor position
    
        If MovingGridX > XSize Then MovingGridX = XSize 'prevent from going past edge of screen

        If MovingGridX < GridX Then MovingGridX = GridX 'if cursor is behind starting point

        PutBox 'redraw entire MusicBox window from SwapScreen
        
        MusicBox.Line (GridX * 16, 0)-(MovingGridX * 16 + 16, MusicBox.Height - 1), vbBlue, B
    
    End If

End Select


End Sub

Private Sub MusicBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'prevents note from being drawn in MouseMove and MouseUp when a file was
'opened by a double-click and cursor was over MusicBox when button was
'released (MouseUp activated with no MouseDown)
If MouseIsDown Then
    MouseIsDown = False
Else
    Exit Sub
End If

Select Case CursorType

'*** Draw Mode ***
Case Is = DrawCursor

    GridX = X \ 16
    
    'finish drawing note
    If Button = 1 Then

        If GridX > XSize Then GridX = XSize 'prevent from going past edge of screen

        channel = CurrentTrack
        If TrackMute(CurrentTrack) = 0 Then StopNote StartY + GridY 'stop playing note that's being placed
        StopNotes

        If GridX < OldGridX Then GridX = OldGridX 'if cursor is behind note's starting point

        'Record to Grid
        Grid(CurrentTrack, StartX + OldGridX) = StartY + GridY 'record starting note
        InstGrid(CurrentTrack, StartX + OldGridX) = CurrentInst(CurrentTrack) 'record starting note's instrument
    
        If GridX > OldGridX Then
            For A = OldGridX + 1 To GridX
                Grid(CurrentTrack, StartX + A) = 1000 + StartY + GridY 'record continuing notes
                InstGrid(CurrentTrack, StartX + A) = CurrentInst(CurrentTrack) 'record continuing notes' instruments
            Next A
        End If
    
        'Check for partial note after new note
        If Grid(CurrentTrack, StartX + GridX + 1) >= 1000 Then
            Grid(CurrentTrack, StartX + GridX + 1) = Grid(CurrentTrack, StartX + GridX + 1) - 1000 'change it to starting note
        End If
    
        DrawGrid

    End If


'*** Select Mode ***
Case Is = SelectCursor

    'finish selecting area
    If Button = 1 Then
    
        GridX = X \ 16
        
        If GridX > XSize Then GridX = XSize 'prevent from going past edge of screen

        If GridX < OldGridX Then GridX = OldGridX 'if cursor is behind note's starting point
    
        SelectEndX = StartX + GridX
        DrawGrid

    End If
    
    
End Select

End Sub


Private Sub DrawGrid()

MusicBox.Cls

'Draw sharp (#) rows different color
For A = 0 To YSize
    If Sharps(A) Then
        MusicBox.Line (0, A * 16)-(MusicBox.Width - 1, A * 16 + 15), &HE0E0E0, BF
    End If
Next A

'Draw columns on MusicBox
For A = 0 To MusicBox.Width - 1 Step 16
    If (StartX + A \ 16) Mod 4 = 0 Then
        Colr = &HFFCC44
    'Else
    '    Colr = &HD0D0D0
    'End If
        MusicBox.Line (A, 0)-(A, MusicBox.Height), Colr
    End If
Next A

'Draw rows on MusicBox
'For A = 0 To MusicBox.Height - 1 Step 16
'    MusicBox.Line (0, A)-(MusicBox.Width, A), &HC0C0C0
'Next A

'draw selected area if visible
If SelectStartX > -1 Then 'there is a selected area
    If Not (SelectStartX > EndX Or SelectEndX < StartX) Then 'at least partially visible
    
        If SelectStartX < StartX Then 'start is off screen
            VisSelStartX = 0
        Else
            VisSelStartX = (SelectStartX - StartX) * 16
        End If
        
        If SelectEndX > EndX Then 'end is off screen
            VisSelEndX = MusicBox.Width - 1
        Else
            VisSelEndX = (SelectEndX - StartX) * 16 + 16
        End If
        
        'draw blue box
        MusicBox.FillStyle = 7
        MusicBox.FillColor = vbBlue
        MusicBox.Line (VisSelStartX, 0)-(VisSelEndX, MusicBox.Height - 1), vbBlue, B
        MusicBox.FillStyle = 1
        
      End If
End If

'draw start position column if visible
If ColumnX >= StartX And ColumnX <= StartX + XSize Then
    A = (ColumnX - StartX) * 16
    MusicBox.FillStyle = 7
    MusicBox.FillColor = vbRed
    MusicBox.Line (A, 0)-(A + 16, MusicBox.Height - 1), vbRed, B
    MusicBox.FillStyle = 1
End If

'Draw notes
For A = 0 To 7
    E = A
    If A >= CurrentTrack Then E = A + 1 'make sure to draw all other tracks before CurrentTrack
    If A = 7 Then E = CurrentTrack 'CurrentTrack is drawn last so it's on the top
    
    If A < 7 Then 'if not drawing top layer then make it dimmed
        Colr = DimmedColrs(E)
    Else
        Colr = TrackColrs(E)
    End If
    
    For B = StartX To StartX + XSize
        C = Grid(E, B)
        If C >= StartY And C <= EndY Then 'draw only notes in viewable area (Y range)
            MusicBox.Line ((B - StartX) * 16 + 2, (C - StartY) * 16 + 2)-((B - StartX) * 16 + 14, (C - StartY) * 16 + 14), Colr, BF
        End If
        
        If C >= 1000 Then
            D = C - 1000
            If D >= StartY And D <= EndY Then
                F = B - StartX
                G = D - StartY
                MusicBox.Line (F * 16 - 2, G * 16 + 2)-(F * 16 + 14, G * 16 + 14), Colr, BF
            End If
        End If
    Next B
    
Next A

End Sub

Sub GetRow(Y%)
Dim RowY As Long

RowY = Y * 16

BitBlt SwapScreen.hDC, 0, 0, MusicBox.Width, 15, MusicBox.hDC, 0, RowY, vbSrcCopy

End Sub

Sub PutRow(Y%)
Dim RowY As Long

RowY = Y * 16

BitBlt MusicBox.hDC, 0, RowY, MusicBox.Width, 15, SwapScreen.hDC, 0, 0, vbSrcCopy

End Sub

Sub GetColumn(X%)
Dim ColX As Long

ColX = X * 16

BitBlt SwapScreen.hDC, 0, 0, 17, MusicBox.Height, MusicBox.hDC, ColX, 0, vbSrcCopy

End Sub

Sub PutColumn(X%)
Dim ColX As Long

ColX = X * 16

BitBlt MusicBox.hDC, ColX, 0, 17, MusicBox.Height, SwapScreen.hDC, 0, 0, vbSrcCopy

End Sub

Sub GetBox()

BitBlt SwapScreen.hDC, 0, 0, MusicBox.Width, MusicBox.Height, MusicBox.hDC, 0, 0, vbSrcCopy

End Sub

Sub PutBox()

BitBlt MusicBox.hDC, 0, 0, MusicBox.Width, MusicBox.Height, SwapScreen.hDC, 0, 0, vbSrcCopy

End Sub

Private Sub PasteBut_Click()
Dim PasteStartX%, PasteEndX%
MusicBox.SetFocus

PasteStartX = ColumnX
PasteEndX = PasteStartX + CopyGridLen

If PasteEndX > 1000 Then PasteEndX = 1000

For A = PasteStartX To PasteEndX
    For B = 0 To 7
        Grid(B, A) = CopyGrid(B, A - PasteStartX)
        InstGrid(B, A) = CopyInstGrid(B, A - PasteStartX)
    Next B
Next A

DrawGrid

End Sub

Private Sub PasteIntoBut_Click()
MusicBox.SetFocus

InsertSpaces CopyGridLen + 1

PasteBut_Click

End Sub

Private Sub InsertColumnBut_Click()
MusicBox.SetFocus

InsertSpaces 1
DrawGrid

End Sub

Private Sub InsertSpaces(NumOfSpaces%)
Dim DestStartX%

DestStartX = ColumnX + NumOfSpaces

'move everything from ColumnX to 1000 over NumOfSpaces spaces
For A = 1000 To DestStartX Step -1
    For B = 0 To 7
        Grid(B, A) = Grid(B, A - NumOfSpaces)
        InstGrid(B, A) = InstGrid(B, A - NumOfSpaces)
    Next B
Next A

'erase grids where the spaces were put
For A = ColumnX To ColumnX + NumOfSpaces - 1
    For B = 0 To 7
        Grid(B, A) = -1
        InstGrid(B, A) = 0
    Next B
Next A

'change start of severed notes from continuing notes to starting notes
For A = 0 To 7
    If Grid(A, DestStartX) >= 1000 Then
        Grid(A, DestStartX) = Grid(A, DestStartX) - 1000
    End If
Next A

End Sub

Private Sub PlayAndStop_Click(Index As Integer)
MusicBox.SetFocus

If Index = 0 Then
    PlayingSong = True
    DisableControls
    PlaySong
Else
    EnableControls
    PlayingSong = False
End If

End Sub

Private Sub PlayAndStop_GotFocus(Index As Integer)
MusicBox.SetFocus

End Sub

Private Sub RewindBut_Click()

MusicBox.SetFocus

StartX = 0
EndX = XSize
ColumnX = 0
DrawGrid

HScroll1.Value = XSizeMid

End Sub


Private Sub TopBarOverlay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim TopBarX%

TopBarX = ScaleX(X, vbTwips, vbPixels) \ 16

ColumnX = StartX + TopBarX
DrawGrid


End Sub

Private Sub TrackMute_Click(Index As Integer)
MusicBox.SetFocus

End Sub

Private Sub TrackSel_Click(Index As Integer)
MusicBox.SetFocus

CurrentTrack = Index

channel = CurrentTrack
ChangeInstrument CurrentInst(CurrentTrack)
Combo1.ListIndex = CurrentInst(CurrentTrack)

TrackColr = TrackColrs(Index)

DrawGrid
End Sub

Private Sub StartNote(Index As Integer)
Flip = 127 - Index 'notes recorded on grid are 127 - midi number

midimsg = &H90 + ((Flip) * &H100) + (volume * &H10000) + channel
midiOutShortMsg hmidi, midimsg

End Sub

Private Sub StopNote(Index As Integer)
Flip = 127 - Index 'notes recorded on grid are 127 - midi number
   
midimsg = &H80 + ((Flip) * &H100) + channel
midiOutShortMsg hmidi, midimsg
   
End Sub

Private Sub StartNotes()

For A = 0 To 7
    If TrackMute(A) = 0 And CurrentTrack <> A Then 'if not muted and not current track
        
        B = (Grid(A, (StartX + NotePlayX))) Mod 1000
        
        If B > 0 Then
            channel = A
            ChangeInstrument InstGrid(A, StartX + NotePlayX) 'change channel's instrument for note being played
            StartNote B
        End If
    End If
Next A

End Sub
Private Sub StopNotes()

For A = 0 To 7
    If TrackMute(A) = 0 And CurrentTrack <> A Then 'if not muted and not current track
    
        B = (Grid(A, (StartX + NotePlayX))) Mod 1000
        
        If B > 0 Then
            channel = A
            ChangeInstrument InstGrid(A, StartX + NotePlayX) 'change channel's instrument for note being played
            StopNote B
        End If
    End If
Next A

End Sub

Private Sub TrackSel_GotFocus(Index As Integer)
MusicBox.SetFocus

End Sub

Private Sub TransposeBut_Click(Index As Integer)
Dim TransStartX%, TransEndX%

MusicBox.SetFocus

'if something is selected, transpose only those notes
'otherwise transpose all notes
If SelectStartX > -1 Then
    TransStartX = SelectStartX
    TransEndX = SelectEndX
Else
    TransStartX = 0
    TransEndX = 1000
End If

'cliked + or - ?
If Index = 0 Then
    C = 1
Else
    C = -1
End If

'transpose notes
For A = 0 To 7
    For B = TransStartX To TransEndX
        If Grid(A, B) > -1 Then
            D = (Grid(A, B) + C)
            
            If D = -1 Then D = 127
            If D = 128 Then D = 0
            If D = 999 Then D = 1127
            If D = 1128 Then D = 1000
            
            Grid(A, B) = D
        End If
    Next B
Next A

DrawGrid

End Sub


Private Sub VScroll1_Change()

If Resizing Then Exit Sub

StartY = VScroll1.Value - YSizeMid
EndY = StartY + YSize
PrintSideBar
DrawGrid
JustScrolled = True 'for when user scrolls window while song is playing

End Sub

Private Sub PrintSideBar()
'Print note names on sidebar
Temp = ""
B = StartY
For A = 0 To YSize
    Temp = Temp & NoteNames(B) & vbCrLf
    If Right(NoteNames(B), 1) = "#" Then 'Find all sharps
        Sharps(A) = True
    Else
        Sharps(A) = False
    End If
    B = B + 1
Next A
SideBar.Caption = Temp
End Sub

Private Sub PrintTopBar()

TopBarOffset = 40

'Print column numbers on topbar
B = 0
For A = StartX To StartX + XSize
    If A Mod 4 = 0 Then
        TopBar(B).Visible = True
        TopBar(B).Left = (A - StartX) * 16 + TopBarOffset
        TopBar(B).Caption = A \ 4
        B = B + 1
    End If
Next A

'make unused labels invisible
For A = B To 18
    TopBar(A).Visible = False
Next A

End Sub

Private Sub VScroll1_Scroll()
VScroll1_Change
End Sub

Private Sub PlaySong()

'if red starting column is not on the screen then move screen
If ColumnX < StartX Or ColumnX > StartX + XSize Then
    StartX = ColumnX
    EndX = StartX + XSize
    If StartX > 999 - XSize Then
        StartX = 999 - XSize 'startx can't be higher than 999 - xsize
        EndX = 999
    End If
End If

PlayX = ColumnX 'start playing from red column
GetColumn StartX
HScroll1.Value = StartX + XSizeMid
HasStarted = False
JustScrolled = False

DrawGrid

'reset oldnote and oldinst variables
For A = 0 To 7
    OldNote(A) = -1
    OldInst(A) = -1
Next A

FindEndOfSong 'so we know when to stop playing

Tick = GetTickCount 'get curent time in milliseconds (I hate the built in VB Timer control!)


'*** Start of song playing loop ***

Do While PlayingSong

    'test how much time left before next note after all routines are done
    'If PlayBarX = 0 Then
    '    InstLabel.Caption = "Extra time = " & (Tick + TempoWait) - GetTickCount & " / " & TempoWait
    'End If

    'loop until 16th note pause has elapsed
    Do While GetTickCount < Tick + TempoWait
        DoEvents
    Loop
    
    Tick = GetTickCount

    PlayNotes
    
    DrawGraphics
    
Loop

'wait another 16th to finish playing last note
Do While GetTickCount < Tick + TempoWait
    DoEvents
Loop

'stop all notes
For A = 0 To 7
    If TrackMute(A) = 0 Then
        channel = A
        If OldNote(A) > -1 Then
            StopNote OldNote(A) Mod 1000
        End If
    End If
Next A

DrawGrid

End Sub

Private Sub PlayNotes()


'Play notes
For Z = 0 To 7
    
    If TrackMute(Z) = 0 Then
    
        channel = Z
        
        'find note types
        If OldNote(Z) = -1 Then OldNoteType = BlankNote
        If OldNote(Z) >= 0 And OldNote(Z) < 1000 Then OldNoteType = StartingNote
        If OldNote(Z) >= 1000 Then OldNoteType = ContinuingNote
        If Grid(Z, PlayX) = -1 Then NoteType = BlankNote
        If Grid(Z, PlayX) >= 0 And Grid(Z, PlayX) < 1000 Then NoteType = StartingNote
        If Grid(Z, PlayX) >= 1000 Then NoteType = ContinuingNote
            
        'change instrument if needed
        'if new note is being played and instrument differs from old instrument
        If NoteType = StartingNote And InstGrid(Z, PlayX) <> OldInst(Z) Then
            ChangeInstrument InstGrid(Z, PlayX)
        End If
        
        'stop old note and start new one
        If OldNoteType <> BlankNote And NoteType = StartingNote Then
            StopNote OldNote(Z) Mod 1000
            StartNote Grid(Z, PlayX) Mod 1000
        End If
        
        'stop old note but don't start new one
        If OldNoteType <> BlankNote And NoteType = BlankNote Then
            StopNote OldNote(Z) Mod 1000
        End If
        
        'start new note but don't stop old
        If OldNoteType = BlankNote And NoteType = StartingNote Then
            StartNote Grid(Z, PlayX) Mod 1000
        End If
        
        'set oldnote to current note's value
        OldNote(Z) = Grid(Z, PlayX)
            
        'set oldinst to current instrument's value if note was started
        If NoteType = StartingNote Then
            OldInst(Z) = InstGrid(Z, PlayX)
        End If
        
    End If
        
Next Z

End Sub

Private Sub DrawGraphics()

'replace background that position bar was drawn over
If PlayBarX < XSize And HasStarted And (Not JustScrolled) Then
    PutColumn PlayBarX
Else
    HasStarted = True
End If
    
JustScrolled = False
    
'show playing position bar
PlayBarX = PlayX - StartX
    
If PlayBarX > XSize Then 'if play bar scrolled off screen move screen over one page
    StartX = StartX + XSize + 1 'move over one page
    EndX = StartX + XSize
    If StartX > 999 - XSize Then
        StartX = 999 - XSize 'prevent showing past end of grid
        EndX = 999
    End If
    PlayBarX = 0 'playbar starts at the left again
    PrintTopBar 'update topbar
    TurningPage = True 'prevents hscroll1_change from calling drawgrid sub
    HScroll1.Value = StartX + XSizeMid 'update horizontal scroll bar
    TurningPage = False 'allow hscroll1_change to call drawgrid sub again
    DrawGrid
    JustScrolled = False 'prevent skipping redraw of leftmost column on next loop
End If
    
GetColumn PlayBarX 'get background which will be behind playbar
        
E = PlayBarX * 16
MusicBox.FillStyle = 7
MusicBox.FillColor = vbGreen
MusicBox.Line (E, 0)-(E + 16, MusicBox.Height - 1), vbGreen, B 'draw green playbar
MusicBox.FillStyle = 1
    
'increase playx (move forward one note)
PlayX = PlayX + 1
    
If PlayX > SongLength Then
    PlayAndStop(1).Value = True
    PlayingSong = False
End If

End Sub

Private Sub ChangeInstrument(Inst As Integer)

midiOutShortMsg hmidi, &HB0 + channel
midiOutShortMsg hmidi, 32 * &H100 + &HB0 + channel
midiOutShortMsg hmidi, Inst * &H100 + &HC0 + channel

End Sub

Private Sub FindEndOfSong()

SongLength = 0

For A = 0 To 7
    For B = 1000 To 0 Step -1
        If Grid(A, B) > -1 Then
            If B > SongLength Then SongLength = B
            Exit For
        End If
    Next B
Next A
                  
End Sub

Private Sub ClearGrids()

'clear grids
For A = 0 To 7
    For B = 0 To 1000
        Grid(A, B) = -1
        InstGrid(A, B) = 0
    Next B
Next A

End Sub

Private Sub DisableControls()

FunctionsFrame.Enabled = False
EditTrackFrame.Enabled = False
MuteTrackFrame.Enabled = False
TransposeFrame.Enabled = False
Combo1.Enabled = False
RewindBut.Enabled = False

End Sub


Private Sub EnableControls()

FunctionsFrame.Enabled = True
EditTrackFrame.Enabled = True
MuteTrackFrame.Enabled = True
TransposeFrame.Enabled = True
Combo1.Enabled = True
RewindBut.Enabled = True

End Sub

