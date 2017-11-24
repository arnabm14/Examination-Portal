VERSION 5.00
Begin VB.Form dForm1 
   Caption         =   "Credentials"
   ClientHeight    =   8340
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   14955
   LinkTopic       =   "Form1"
   ScaleHeight     =   8340
   ScaleWidth      =   14955
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3960
      TabIndex        =   12
      Top             =   6720
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9240
      TabIndex        =   11
      Top             =   6720
      Width           =   2295
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   6600
      TabIndex        =   4
      Top             =   5520
      Width           =   4935
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   6600
      TabIndex        =   3
      Top             =   4560
      Width           =   4935
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   6600
      TabIndex        =   2
      Top             =   3600
      Width           =   4935
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   6600
      TabIndex        =   1
      Top             =   2640
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   6600
      TabIndex        =   0
      Top             =   1680
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "WELCOME TO THE TEST"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4320
      TabIndex        =   5
      Top             =   360
      Width           =   7095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "NAME"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   6
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "STREAM"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "YEAR"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   8
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "SECTION"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   9
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "ROLL NO"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   10
      Top             =   5640
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   8415
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14895
   End
End
Attribute VB_Name = "dForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, b, c, d, e, f, g, h


Private Sub Command1_Click()
a = MsgBox("Are you sure you want to start?", vbOKCancel, "Start")
If a = 1 Then
Unload dForm2
dForm2.Show
dForm1.Visible = False
End If
End Sub

Private Sub Command2_Click()
a = MsgBox("Are you sure you want to cancel?", vbOKCancel, "Quit")
If a = 1 Then
dForm1.Visible = False
Unload Form1
Form1.Show
End If

End Sub

Private Sub Form_Load()
Image1.Picture = LoadPicture(App.Path & "\2.jpg")
End Sub

