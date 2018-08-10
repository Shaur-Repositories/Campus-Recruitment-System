VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00FFFF00&
   Caption         =   "Apti Section 1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Adobe Myungjo Std M"
      Size            =   18
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form5"
   Picture         =   "Form5.frx":0000
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Left            =   18240
      Top             =   1560
   End
   Begin VB.TextBox Text6 
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
      Left            =   17040
      TabIndex        =   14
      Top             =   6720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Left            =   12240
      TabIndex        =   13
      Top             =   8520
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Left            =   12240
      TabIndex        =   12
      Top             =   6840
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Left            =   12240
      TabIndex        =   11
      Top             =   5280
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Left            =   12240
      TabIndex        =   10
      Top             =   3960
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Left            =   12240
      TabIndex        =   9
      Top             =   2400
      Width           =   1575
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   1320
      Picture         =   "Form5.frx":4C44D
      ScaleHeight     =   945
      ScaleWidth      =   10305
      TabIndex        =   7
      Top             =   5040
      Width           =   10335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF00FF&
      Caption         =   "NEXT >>>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   16440
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9240
      Width           =   2655
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   1320
      Picture         =   "Form5.frx":50F05
      ScaleHeight     =   1425
      ScaleWidth      =   10305
      TabIndex        =   4
      Top             =   8040
      Width           =   10335
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   1320
      Picture         =   "Form5.frx":571BD
      ScaleHeight     =   1545
      ScaleWidth      =   10305
      TabIndex        =   3
      Top             =   6240
      Width           =   10335
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   1320
      Picture         =   "Form5.frx":5D0E3
      ScaleHeight     =   1185
      ScaleWidth      =   10305
      TabIndex        =   1
      Top             =   3600
      Width           =   10335
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   1320
      Picture         =   "Form5.frx":63FF9
      ScaleHeight     =   1545
      ScaleWidth      =   10305
      TabIndex        =   0
      Top             =   1800
      Width           =   10335
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "2."
      ForeColor       =   &H00FF00FF&
      Height          =   615
      Left            =   480
      TabIndex        =   22
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "3."
      ForeColor       =   &H00FF80FF&
      Height          =   615
      Left            =   480
      TabIndex        =   21
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "4."
      ForeColor       =   &H00FF80FF&
      Height          =   615
      Left            =   480
      TabIndex        =   20
      Top             =   6720
      Width           =   615
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "5."
      ForeColor       =   &H00FF80FF&
      Height          =   615
      Left            =   480
      TabIndex        =   19
      Top             =   8520
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1."
      ForeColor       =   &H00FF80FF&
      Height          =   615
      Left            =   480
      TabIndex        =   18
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Section 1"
      BeginProperty Font 
         Name            =   "Adobe Myungjo Std M"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   5760
      TabIndex        =   17
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "1200"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   17880
      TabIndex        =   16
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "timer:"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   16440
      TabIndex        =   15
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "your answers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   11400
      TabIndex        =   8
      Top             =   1320
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "(1)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   10560
      TabIndex        =   6
      Top             =   9960
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "        aptitude  test"
      BeginProperty Font 
         Name            =   "Trajan Pro"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   2040
      TabIndex        =   2
      Top             =   120
      Width           =   10215
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a, r As Integer
a = 0
If Text1.Text = "c" Then
a = a + 3
ElseIf Text1.Text = "" Then
a = a
Else
a = a - 1
End If
If Text2.Text = "c" Then
a = a + 3
ElseIf Text2.Text = "" Then
a = a
Else
a = a - 1
End If
If Text3.Text = "7" Then
a = a + 3
ElseIf Text3.Text = "" Then
a = a
Else
a = a - 1
End If
If Text4.Text = "a" Then
a = a + 3
ElseIf Text4.Text = "" Then
a = a
Else
a = a - 1
End If
If Text5.Text = "0.31" Then
a = a + 3
ElseIf Text5.Text = "" Then
a = a
Else
a = a - 1
End If
Text6.Text = a
r = MsgBox("are u sure", vbCritical + vbYesNo, "Shaur industries....?")
If r = vbYes Then
Form3.Show
Form5.Hide
End If

End Sub

Private Sub Form_Load()
Timer1.Interval = 1000
End Sub

Private Sub Timer1_Timer()
If Label6.Caption = 0 Then
Timer1.Enabled = False
MsgBox "test is over", vbvbOKOnly, "Shaur industries"
Label4.Visible = True
Text1.Locked = True
Text2.Locked = True
Text3.Locked = True
Text4.Locked = True
Text5.Locked = True
Else
Label6.Caption = Label6.Caption - 1
End If
End Sub
