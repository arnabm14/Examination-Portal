VERSION 5.00
Begin VB.Form aForm1 
   Caption         =   "Admin"
   ClientHeight    =   5610
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9795
   LinkTopic       =   "Form1"
   ScaleHeight     =   5600
   ScaleMode       =   0  'User
   ScaleWidth      =   9746.269
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "GO BACK"
      Height          =   495
      Left            =   8040
      TabIndex        =   2
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete Question(s)"
      Height          =   495
      Left            =   4200
      TabIndex        =   1
      Top             =   3840
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Question(s)"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome Admin, Choose one option."
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   9735
   End
   Begin VB.Image Image1 
      Height          =   5655
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9855
   End
End
Attribute VB_Name = "aForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
aForm1.Hide
Unload aForm2
aForm2.Show
End Sub


Private Sub Command2_Click()
aForm1.Hide
Unload aForm4
aForm4.Show
End Sub

Private Sub Command3_Click()
aForm1.Hide
Unload Form1
Form1.Show
End Sub

Private Sub Form_Load()
Image1.Picture = LoadPicture(App.Path & "\3.jpg")
End Sub
