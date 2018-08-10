VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form8 
   Caption         =   "eligble"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form8"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   800
      Left            =   1080
      Top             =   3120
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Gloucester MT Extra Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8760
      Width           =   3375
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp1 
      Height          =   735
      Left            =   1440
      TabIndex        =   1
      Top             =   5760
      Visible         =   0   'False
      Width           =   3855
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
      _cx             =   6800
      _cy             =   1296
   End
   Begin VB.Image Image4 
      Height          =   11520
      Left            =   0
      Picture         =   "Form8.frx":0000
      Top             =   0
      Width           =   20490
   End
   Begin VB.Image Image3 
      Height          =   11520
      Left            =   0
      Picture         =   "Form8.frx":202A8
      Top             =   0
      Width           =   20490
   End
   Begin VB.Image Image2 
      Height          =   11520
      Left            =   0
      Picture         =   "Form8.frx":6C675
      Top             =   0
      Width           =   20490
   End
   Begin VB.Image Image1 
      Height          =   11520
      Left            =   0
      Picture         =   "Form8.frx":90564
      Top             =   0
      Width           =   20490
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
wmp1.URL = "C:\MEDIA\sounds\5.wav"
End
End Sub

Private Sub Timer1_Timer()
If Image1.Visible = True Then
Image2.Visible = True
Image1.Visible = False
ElseIf Image2.Visible = True Then
Image2.Visible = False
Image3.Visible = True
ElseIf Image3.Visible = True Then
Image3.Visible = False
Image4.Visible = True
ElseIf Image4.Visible = True Then
Image4.Visible = False
Image1.Visible = True
End If
End Sub
