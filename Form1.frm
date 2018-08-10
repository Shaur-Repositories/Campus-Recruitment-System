VERSION 5.00
Object = "{28D47522-CF84-11D1-834C-00A0249F0C28}#1.0#0"; "gif89 .dll"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   BackColor       =   &H000080FF&
   Caption         =   "Shaur Industries"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   17.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   376.634
   ScaleMode       =   0  'User
   ScaleWidth      =   162.133
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   1680
      Top             =   2520
   End
   Begin GIF89LibCtl.Gif89a Gif89a1 
      Height          =   10095
      Left            =   0
      OleObjectBlob   =   "Form1.frx":38AAB6
      TabIndex        =   1
      Top             =   0
      Width           =   19455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "click to enter"
      DisabledPicture =   "Form1.frx":38AAF8
      DownPicture     =   "Form1.frx":3CCE00
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8640
      MaskColor       =   &H0000FFFF&
      Picture         =   "Form1.frx":3E63CA
      TabIndex        =   0
      Top             =   8160
      Width           =   3015
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp1 
      Height          =   1335
      Left            =   2760
      TabIndex        =   2
      Top             =   2880
      Visible         =   0   'False
      Width           =   6855
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
      _cx             =   12091
      _cy             =   2355
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
wmp1.URL = "C:\MEDIA\sounds\4.mp3"
Gif89a1.FileName = "c:\media\v3.gif"
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
Form10.Show
Form1.Hide
Timer1.Enabled = False
End Sub
