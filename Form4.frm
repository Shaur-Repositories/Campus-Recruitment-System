VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form4 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H000000FF&
   Caption         =   "Apti Rules"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   ForeColor       =   &H0000FFFF&
   LinkTopic       =   "Form4"
   Picture         =   "Form4.frx":0000
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   120
      Left            =   13320
      Top             =   7680
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   615
      Left            =   6000
      TabIndex        =   2
      Top             =   5520
      Visible         =   0   'False
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   1085
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   11640
      Top             =   2400
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Logout"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   17520
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9720
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Start Test"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   15360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9720
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp1 
      Height          =   495
      Left            =   7560
      TabIndex        =   5
      Top             =   1320
      Visible         =   0   'False
      Width           =   30
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
      _cx             =   53
      _cy             =   873
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Loading...."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   735
      Left            =   8640
      TabIndex        =   4
      Top             =   6960
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H00C000C0&
      Height          =   615
      Left            =   8280
      TabIndex        =   3
      Top             =   6240
      Visible         =   0   'False
      Width           =   3735
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim r As Integer
Option Explicit
Private StrCap As String
Private Num As Long
Private Sub Form_Load()
StrCap = Label2.Caption
End Sub


Private Sub Command1_Click()
wmp1.URL = "C:\MEDIA\sounds\5.wav"
'Command1.Enabled = False
Command2.Enabled = False
Timer1.Enabled = True
Timer2.Enabled = True
ProgressBar1.Visible = True
Label1.Visible = True
Label2.Visible = True
wmp1.URL = "C:\MEDIA\sounds\7.mp3"
End Sub

Private Sub Command2_Click()
wmp1.URL = "C:\MEDIA\sounds\5.wav"
r = MsgBox("are u sure", vbCritical + vbYesNo, "Shaur industries....?")
If r = vbYes Then
Form2.Show
Form4.Hide
End If
End Sub


Private Sub Timer1_Timer()

Dim i As Integer
i = ProgressBar1.Value
Timer1.Interval = Rnd * 300 + 10
ProgressBar1.Value = i + 4
Label1.Caption = ProgressBar1.Value & "%"
If Label1.Caption = 100 & "%" Then
ProgressBar1.Visible = False
Label1.Visible = False
Label2.Visible = False
Form5.Show
Unload Me
End If
End Sub

Private Sub Timer2_Timer()
If Label2.Caption <> StrCap Then
Label2.Alignment = 0
Label2.Caption = Left(StrCap, Len(Label2.Caption) + 1)
Else
Label2.Caption = ""
Num = 0
End If
End Sub

