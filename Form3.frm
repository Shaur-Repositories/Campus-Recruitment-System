VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form3 
   BackColor       =   &H00FFFF00&
   Caption         =   "Apti Section 2"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form3"
   Picture         =   "Form3.frx":0000
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   735
      Left            =   5880
      TabIndex        =   21
      Top             =   5400
      Visible         =   0   'False
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   1296
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   120
      Left            =   14040
      Top             =   4320
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   13920
      Top             =   5280
   End
   Begin VB.Timer Timer1 
      Left            =   18240
      Top             =   2280
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   17760
      TabIndex        =   12
      Top             =   7920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11280
      TabIndex        =   11
      Top             =   8880
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11280
      TabIndex        =   10
      Top             =   7320
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11280
      TabIndex        =   9
      Top             =   5520
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11280
      TabIndex        =   8
      Top             =   3600
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11280
      TabIndex        =   7
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   17400
      TabIndex        =   6
      Top             =   9360
      Width           =   2295
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   1320
      Picture         =   "Form3.frx":18E2E
      ScaleHeight     =   1185
      ScaleWidth      =   9345
      TabIndex        =   4
      Top             =   8640
      Width           =   9375
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   1320
      Picture         =   "Form3.frx":1D928
      ScaleHeight     =   1665
      ScaleWidth      =   9345
      TabIndex        =   3
      Top             =   6840
      Width           =   9375
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   1320
      Picture         =   "Form3.frx":24B25
      ScaleHeight     =   1785
      ScaleWidth      =   9345
      TabIndex        =   2
      Top             =   4920
      Width           =   9375
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   1320
      Picture         =   "Form3.frx":2BD01
      ScaleHeight     =   1665
      ScaleWidth      =   9345
      TabIndex        =   1
      Top             =   3120
      Width           =   9375
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   1320
      Picture         =   "Form3.frx":34792
      ScaleHeight     =   2145
      ScaleWidth      =   9345
      TabIndex        =   0
      Top             =   840
      Width           =   9375
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp1 
      Height          =   375
      Left            =   16800
      TabIndex        =   24
      Top             =   3840
      Visible         =   0   'False
      Width           =   2175
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
      _cx             =   3836
      _cy             =   661
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Analising Result....."
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   15000
      TabIndex        =   23
      Top             =   6240
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   15120
      TabIndex        =   22
      Top             =   5520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "2."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   20
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "3."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   19
      Top             =   5520
      Width           =   495
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "4."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   18
      Top             =   7320
      Width           =   495
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "5."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   17
      Top             =   8880
      Width           =   495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "1."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   16
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Section II"
      BeginProperty Font 
         Name            =   "Berlin Sans FB"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   15
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "1200"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   18120
      TabIndex        =   14
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "timer :"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   16920
      TabIndex        =   13
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      Caption         =   "(2)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10320
      TabIndex        =   5
      Top             =   9960
      Width           =   615
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, r As Integer
Option Explicit
Private StrCap As String
Private Num As Long
Private Sub Form_Load()
Timer1.Interval = 1000
StrCap = Label11.Caption
End Sub

Private Sub Command1_Click()

a = 0
If Text1.Text = "a" Then
a = a + 3
ElseIf Text1.Text = "" Then
a = a
Else
a = a - 1
End If
If Text2.Text = "93.3" Then
a = a + 3
ElseIf Text2.Text = "" Then
a = a
Else
a = a - 1
End If
If Text3.Text = "8" Then
a = a + 3
ElseIf Text3.Text = "" Then
a = a
Else
a = a - 1
End If
If Text4.Text = "c" Then
a = a + 3
ElseIf Text4.Text = "" Then
a = a
Else
a = a - 1
End If
If Text5.Text = "c" Then
a = a + 3
ElseIf Text5.Text = "" Then
a = a
Else
a = a - 1
End If
Text6.Text = a
r = MsgBox("are u sure ?", vbCritical + vbYesNo, "Shaur industries....?")
If r = vbYes Then
Timer1.Enabled = False
Text1.Locked = True
Text2.Locked = True
Text3.Locked = True
Text4.Locked = True
Text5.Locked = True
Form6.Label3.Caption = Form5.Text6.Text
Form6.Label4.Caption = Form3.Text6.Text
Timer2.Enabled = True
Timer3.Enabled = True
ProgressBar1.Visible = True
Label10.Visible = True
Label11.Visible = True
wmp1.URL = "C:\MEDIA\sounds\7.mp3"
End If
End Sub



Private Sub Timer1_Timer()
If Label3.Caption = 0 Then
Timer1.Enabled = False
MsgBox "test is over"
Text1.Locked = True
Text2.Locked = True
Text3.Locked = True
Text4.Locked = True
Text5.Locked = True
Else
Label3.Caption = Label3.Caption - 1
End If
End Sub
Private Sub Timer2_Timer()
Dim i As Integer
i = ProgressBar1.Value
Timer2.Interval = Rnd * 300 + 10
ProgressBar1.Value = i + 2
Label10.Caption = ProgressBar1.Value & "%"
If Label10.Caption = 100 & "%" Then
ProgressBar1.Visible = False
Label10.Visible = False
Label11.Visible = False
Form6.Show
Unload Me
End If
End Sub

Private Sub Timer3_Timer()
If Label11.Caption <> StrCap Then
Label11.Alignment = 0
Label11.Caption = Left(StrCap, Len(Label11.Caption) + 1)
Else
Label11.Caption = ""
Num = 0
End If
End Sub
