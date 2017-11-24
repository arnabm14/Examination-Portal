VERSION 5.00
Begin VB.Form aForm2 
   Caption         =   "Admin"
   ClientHeight    =   5610
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9795
   LinkTopic       =   "Form2"
   ScaleHeight     =   5570.212
   ScaleMode       =   0  'User
   ScaleWidth      =   9760.837
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text7 
      Height          =   495
      Left            =   600
      TabIndex        =   13
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   4080
      TabIndex        =   11
      Top             =   3240
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   8040
      TabIndex        =   6
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   6240
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "GO BACK"
      Height          =   495
      Left            =   7440
      TabIndex        =   2
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add "
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   4200
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   1695
      Left            =   2280
      TabIndex        =   0
      Text            =   "Add your Question here ."
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Q. No"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   14
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Option-1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Option-2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Correct Option"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   12
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Option-4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   10
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Option-3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   9
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   5655
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9855
   End
End
Attribute VB_Name = "aForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim conec As New ADODB.Connection
    Dim rs As New ADODB.Recordset
Sub clear()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    End Sub
Private Sub Command1_Click()
    conec.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " & App.Path & "\Database.mdb; Persist Security Info=False"
    conec.CursorLocation = adUseClient
    conec.Open
    rs.Open "select * from Table1", conec, adOpenDynamic, adLockPessimistic
    rs.AddNew
    rs.Fields("Question").Value = Text1.Text
    rs.Fields("First_option").Value = Text2.Text
    rs.Fields("Second_option").Value = Text3.Text
    rs.Fields("Third_option").Value = Text4.Text
    rs.Fields("Fourth_option").Value = Text5.Text
    rs.Fields("Correct_option").Value = Text6.Text
    rs.Fields("Qno").Value = Val(Text7.Text)
    MsgBox "DATA RECORDED"
    rs.Update
    conec.Close
    Unload aForm2
    aForm1.Show
End Sub

Private Sub Command2_Click()
aForm2.Hide
aForm1.Show

End Sub

Private Sub Form_Load()
Image1.Picture = LoadPicture(App.Path & "\3.jpg")

End Sub
