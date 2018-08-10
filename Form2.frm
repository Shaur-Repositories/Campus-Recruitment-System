VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form2 
   Caption         =   "Student Login page"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\SHAUR\Documents\vb project\reg.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "stud"
      Top             =   6840
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0000FFFF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Gloucester MT Extra Condensed"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   15840
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9120
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Gloucester MT Extra Condensed"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10080
      TabIndex        =   6
      Top             =   6600
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "Register"
      BeginProperty Font 
         Name            =   "Gloucester MT Extra Condensed"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   15840
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   960
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "Sign In"
      BeginProperty Font 
         Name            =   "Gloucester MT Extra Condensed"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6600
      MaskColor       =   &H0000FFFF&
      Picture         =   "Form2.frx":40771
      TabIndex        =   4
      Top             =   6600
      UseMaskColor    =   -1  'True
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "20th Century Font"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   9840
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   4440
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "20th Century Font"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      HideSelection   =   0   'False
      Left            =   9840
      MouseIcon       =   "Form2.frx":7B2D58
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   3360
      Width           =   2055
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp1 
      Height          =   255
      Left            =   9600
      TabIndex        =   8
      Top             =   1440
      Visible         =   0   'False
      Width           =   3615
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
      _cx             =   6376
      _cy             =   450
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   375
      Left            =   7920
      TabIndex        =   1
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "User I.D."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   375
      Left            =   7920
      TabIndex        =   0
      Top             =   3360
      Width           =   1815
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim r As Integer
Private Sub Command1_Click()
wmp1.URL = "C:\MEDIA\sounds\5.wav"
Dim strName As String
Dim strPass As String
Dim pesan As String
Data1.Refresh
strName = Text1.Text
strPass = Text2.Text
Do Until Data1.Recordset.EOF
If Data1.Recordset.Fields("id").Value = strName And Data1.Recordset.Fields("pass").Value = strPass Then
MsgBox "logged in", vbOKOnly, "Shaur Industries"
Form4.Show
Form2.Hide
Exit Sub
Else
Data1.Recordset.MoveNext
End If
Loop
pesan = MsgBox("Invalid id or password, try again!", vbOKOnly, "Shaur Industries...?")
If (pesan = 1) Then
Form2.Show
End If
End Sub

Private Sub Command2_Click()
wmp1.URL = "C:\MEDIA\sounds\5.wav"
form7.Data1.Recordset.AddNew
form7.Show
Form2.Hide
End Sub

Private Sub Command3_Click()
wmp1.URL = "C:\MEDIA\sounds\5.wav"
Text1.Text = ""
Text2.Text = ""
End Sub

Private Sub Command4_Click()
wmp1.URL = "C:\MEDIA\sounds\5.wav"
End
End Sub

