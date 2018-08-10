VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form form7 
   AutoRedraw      =   -1  'True
   Caption         =   "regestration"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form7"
   Picture         =   "Form7.frx":0000
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   3480
      TabIndex        =   47
      Top             =   3000
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Minion Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   128516097
      CurrentDate     =   42687
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H0000FF00&
      Caption         =   "<<<log in"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   17760
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   720
      Width           =   1815
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   90
      Left            =   8280
      Top             =   6840
   End
   Begin VB.TextBox Text13 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10800
      TabIndex        =   41
      Text            =   "Text13"
      Top             =   3960
      Width           =   2775
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7560
      Top             =   6720
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FF00FF&
      Caption         =   "Register"
      BeginProperty Font 
         Name            =   "Bernard MT Condensed"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   8880
      Visible         =   0   'False
      Width           =   2655
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   735
      Left            =   5760
      TabIndex        =   37
      Top             =   5520
      Visible         =   0   'False
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   1296
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog5 
      Left            =   8640
      Top             =   9960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog4 
      Left            =   11760
      Top             =   9960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog3 
      Left            =   9960
      Top             =   9960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   9240
      Top             =   9960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10560
      Top             =   9960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog6 
      Left            =   11160
      Top             =   9960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Bitmaps (*.bmp)|*.bmp|"
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H000000FF&
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "Bernard MT Condensed"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   15240
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   8880
      Width           =   2535
   End
   Begin VB.TextBox Text12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   35
      Text            =   "Text12"
      Top             =   9240
      Width           =   1095
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   34
      Text            =   "Text11"
      Top             =   8520
      Width           =   1095
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   33
      Text            =   "Text10"
      Top             =   7800
      Width           =   1095
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   32
      Text            =   "Text9"
      Top             =   7080
      Width           =   1095
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   31
      Text            =   "Text8"
      Top             =   3960
      Width           =   2655
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10800
      TabIndex        =   30
      Text            =   "Text7"
      Top             =   1440
      Width           =   5415
   End
   Begin VB.PictureBox Picture5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      ScaleHeight     =   435
      ScaleWidth      =   915
      TabIndex        =   29
      Top             =   9240
      Width           =   975
   End
   Begin VB.PictureBox Picture4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      ScaleHeight     =   435
      ScaleWidth      =   915
      TabIndex        =   28
      Top             =   8520
      Width           =   975
   End
   Begin VB.PictureBox Picture3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      ScaleHeight     =   435
      ScaleWidth      =   915
      TabIndex        =   27
      Top             =   7800
      Width           =   975
   End
   Begin VB.PictureBox Picture2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      ScaleHeight     =   435
      ScaleWidth      =   915
      TabIndex        =   26
      Top             =   7080
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      ScaleHeight     =   435
      ScaleWidth      =   915
      TabIndex        =   25
      Top             =   6360
      Width           =   975
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   24
      Text            =   "Text6"
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00800080&
      Caption         =   "upload"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   9240
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0000FF00&
      Caption         =   "upload"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   8520
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H008080FF&
      Caption         =   "upload"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000FFFF&
      Caption         =   "upload"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFF00&
      Caption         =   "upload"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "Bernard MT Condensed"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11760
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8880
      Width           =   2655
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\SHAUR\Documents\vb project\reg.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "stud"
      Top             =   8520
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      DataField       =   "pass"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10800
      TabIndex        =   12
      Text            =   "Text5"
      Top             =   3240
      Width           =   2775
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      DataField       =   "id"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10800
      TabIndex        =   11
      Text            =   "Text4"
      Top             =   2520
      Width           =   4215
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   9
      Text            =   "Text2"
      Top             =   2280
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   10
      Text            =   "Text3"
      Top             =   3000
      Width           =   2655
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp1 
      Height          =   495
      Left            =   1320
      TabIndex        =   48
      Top             =   840
      Visible         =   0   'False
      Width           =   3255
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   5741
      _cy             =   873
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(will be used as your user i.d.)"
      BeginProperty Font 
         Name            =   "20th Century Font"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   15240
      TabIndex        =   46
      Top             =   2640
      Width           =   2940
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "*** Please write your percentage  and  uload marksheet of the following year ***"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   45
      Top             =   10080
      Width           =   5655
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Analising informatation and documents......."
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   615
      Left            =   12120
      TabIndex        =   43
      Top             =   6960
      Visible         =   0   'False
      Width           =   6495
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "<Date/Month/Year>"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   42
      Top             =   3600
      Width           =   2535
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Re type passward"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   7320
      TabIndex        =   40
      Top             =   3960
      Width           =   3570
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1335
      Left            =   9600
      TabIndex        =   38
      Top             =   6480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "B.E percentage"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   18
      Top             =   9240
      Width           =   1935
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "T.E percentage"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   17
      Top             =   8520
      Width           =   2055
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "S.E percentage"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   16
      Top             =   7800
      Width           =   1935
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "F.E pecentage"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   15
      Top             =   7080
      Width           =   2175
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Passward"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7320
      TabIndex        =   13
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "12th percentage"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   6360
      Width           =   2415
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7320
      TabIndex        =   6
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      TabIndex        =   5
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Full Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   7320
      TabIndex        =   4
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "D.O.B."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Father's name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Student Regestration"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5040
      TabIndex        =   0
      Top             =   240
      Width           =   6615
   End
End
Attribute VB_Name = "form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, d, e, f, g As Integer
Option Explicit
Private StrCap As String
Private Num As Long



Private Sub Command9_Click()
wmp1.URL = "C:\MEDIA\sounds\5.wav"
form7.Hide
Form2.Show
End Sub

Private Sub Form_Load()
StrCap = Label17.Caption
End Sub
Private Sub Command1_Click()
wmp1.URL = "C:\MEDIA\sounds\5.wav"
If Text1.Text = "" Then
MsgBox "enter name", vbYes, "Shaur Industries"
End If
If Text2.Text = "" Then
MsgBox "enter father's name", vbYes, "Shaur Industries"
End If
'If Text3.Text = "" Then
'MsgBox "enter your D.O.B", vbYes, "Shaur Industries"
'End If
If Text4.Text = "" Then
MsgBox "enter email id", vbYes, "Shaur Industries"
End If
If Text5.Text = "" Then
MsgBox "passward can't be empty ", vbYes, "Shaur Industries"
End If
If Text6.Text = "" Then
MsgBox " 12th marks required", vbYes, "Shaur Industries"
End If
If Text7.Text = "" Then
MsgBox "address required", vbYes, "Shaur Industries"
End If
If Text8.Text = "" Then
MsgBox "enter mobile no.", vbYes, "Shaur Industries"
End If
If Text9.Text = "" Then
MsgBox "F.E. marks required", vbYes, "Shaur Industries"
End If
If Text10.Text = "" Then
MsgBox "S.E. marks required", vbYes, "Shaur Industries"
End If
If Text11.Text = "" Then
MsgBox "T.E. marks required", vbYes, "Shaur Industries"
End If
If Text12.Text = "" Then
MsgBox "B.E. marks required", vbYes, "Shaur Industries"
End If
If Text13.Text = "" Then
MsgBox "Re enter the passward", vbYes, "Shaur Industries"
End If
If Picture1.Picture = Empty Or Picture2.Picture = Empty Or Picture3.Picture = Empty Or Picture4.Picture = Empty Or Picture5.Picture = Empty Then
MsgBox "upload all documents", vbYes, "Shaur Industries"
End If
If Text1.Text <> "" And Text2.Text <> "" And Text4.Text <> "" And Text5.Text <> "" And Text6.Text <> "" And Text7.Text <> "" And Text8.Text <> "" And Text9.Text <> "" And Text10.Text <> "" And Text11.Text <> "" And Text12.Text <> "" And Picture1.Picture <> Empty And Picture2.Picture <> Empty And Picture3.Picture <> Empty And Picture4.Picture <> Empty And Picture5.Picture <> Empty Then
If Text13.Text <> Text5.Text Then
MsgBox "re-enter passward correctly", vbYes, "Shaur Industries"
Text13.Text = ""
Else
Command7.Visible = False
Command8.Visible = True
Command1.Visible = False
End If
End If
'If a >= 60 And d >= 50 And e >= 50 And f >= 50 And g >= 50 Then
'Data1.Recordset.Update
'MsgBox "successful", vbYes, "Shaur Industries"
'Else
'MsgBox "fuck", vbYes, "Shaur Industries"
'End If
End Sub

Private Sub Command2_Click()
wmp1.URL = "C:\MEDIA\sounds\5.wav"
CommonDialog1.Filter = "graphic files *.bmp;*.gif;*.jpg"
CommonDialog1.ShowOpen
Picture1.Picture = LoadPicture(CommonDialog1.FileName)
End Sub
Private Sub Command3_Click()
wmp1.URL = "C:\MEDIA\sounds\5.wav"
CommonDialog2.Filter = "graphic files *.bmp;*.gif;*.jpg"
CommonDialog2.ShowOpen
Picture2.Picture = LoadPicture(CommonDialog2.FileName)
End Sub

Private Sub Command4_Click()
wmp1.URL = "C:\MEDIA\sounds\5.wav"
CommonDialog3.Filter = "graphic files *.bmp;*.gif;*.jpg"
CommonDialog3.ShowOpen
Picture3.Picture = LoadPicture(CommonDialog3.FileName)
End Sub
Private Sub Command5_Click()
wmp1.URL = "C:\MEDIA\sounds\5.wav"
CommonDialog4.Filter = "graphic files *.bmp;*.gif;*.jpg"
CommonDialog4.ShowOpen
Picture4.Picture = LoadPicture(CommonDialog4.FileName)
End Sub
Private Sub Command6_Click()
wmp1.URL = "C:\MEDIA\sounds\5.wav"
CommonDialog5.Filter = "graphic files *.bmp;*.gif;*.jpg"
CommonDialog5.ShowOpen
Picture5.Picture = LoadPicture(CommonDialog5.FileName)
End Sub

Private Sub Command7_Click()
wmp1.URL = "C:\MEDIA\sounds\5.wav"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Set Picture5.Picture = LoadPicture
Set Picture4.Picture = LoadPicture
Set Picture3.Picture = LoadPicture
Set Picture2.Picture = LoadPicture
Set Picture1.Picture = LoadPicture
Data1.Recordset.AddNew
End Sub

Private Sub Command8_Click()
wmp1.URL = "C:\MEDIA\sounds\5.wav"
Text1.Locked = True
Text2.Locked = True
Text3.Locked = True
Text4.Locked = True
Text5.Locked = True
Text6.Locked = True
Text7.Locked = True
Text8.Locked = True
Text9.Locked = True
Text10.Locked = True
Text11.Locked = True
Text12.Locked = True
Text13.Locked = True
Timer1.Enabled = True
Timer2.Enabled = True
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
DTPicker1.Enabled = False
ProgressBar1.Visible = True
Label14.Visible = True
Label17.Visible = True
'Command8.Enabled = False
Command9.Enabled = False
wmp1.URL = "C:\MEDIA\sounds\7.mp3"
End Sub


Private Sub Timer1_Timer()

Dim i As Integer
i = ProgressBar1.Value
Timer1.Interval = Rnd * 300 + 10
ProgressBar1.Value = i + 1
Label14.Caption = ProgressBar1.Value & "%"
If Label14.Caption = 100 & "%" Then
Data1.Recordset.AddNew
a = Text6.Text
d = Text9.Text
e = Text10.Text
f = Text11.Text
g = Text12.Text
If a >= 60 And d >= 50 And e >= 50 And f >= 50 And g >= 50 Then
Data1.Recordset.Update
form7.Hide
Form8.Show
Else
Data1.Recordset.MoveLast
Data1.Recordset.Delete
form7.Hide
Form9.Show
End If
ProgressBar1.Visible = False
Label14.Visible = False
Label17.Visible = False
Unload Me
End If
End Sub
Private Sub Timer2_Timer()
If Label17.Caption <> StrCap Then
Label17.Alignment = 0
Label17.Caption = Left(StrCap, Len(Label17.Caption) + 1)
Else
Label17.Caption = ""
Num = 0
End If
End Sub
