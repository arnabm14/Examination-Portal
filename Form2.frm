VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form2 
   Caption         =   "Sign Up"
   ClientHeight    =   10995
   ClientLeft      =   390
   ClientTop       =   405
   ClientWidth     =   19875
   LinkTopic       =   "Form2"
   ScaleHeight     =   10995
   ScaleWidth      =   19875
   Begin VB.CommandButton Command2 
      Caption         =   "SIGN UP"
      Height          =   975
      Left            =   12840
      TabIndex        =   21
      Top             =   8640
      Width           =   4695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form2.frx":0000
      Left            =   15240
      List            =   "Form2.frx":000A
      TabIndex        =   20
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SUBMIT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   13920
      TabIndex        =   18
      Top             =   11880
      Width           =   4095
   End
   Begin VB.TextBox Text6 
      Height          =   855
      Left            =   6240
      TabIndex        =   17
      Top             =   11160
      Width           =   4335
   End
   Begin VB.TextBox Text5 
      Height          =   855
      Left            =   15240
      TabIndex        =   16
      Top             =   7080
      Width           =   4335
   End
   Begin VB.TextBox Text4 
      Height          =   855
      Left            =   6240
      TabIndex        =   15
      Top             =   8640
      Width           =   4335
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   735
      Left            =   6360
      TabIndex        =   14
      Top             =   7440
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1296
      _Version        =   393216
      Format          =   82116609
      CurrentDate     =   42849
   End
   Begin VB.OptionButton Option2 
      Caption         =   "FEMALE"
      Height          =   615
      Left            =   8280
      TabIndex        =   13
      Top             =   6240
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "MALE"
      Height          =   855
      Left            =   6480
      TabIndex        =   12
      Top             =   6120
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   855
      IMEMode         =   3  'DISABLE
      Left            =   6240
      PasswordChar    =   "*"
      TabIndex        =   11
      Top             =   4680
      Width           =   4335
   End
   Begin VB.TextBox Text2 
      Height          =   855
      Left            =   6240
      TabIndex        =   10
      Top             =   3360
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   6120
      TabIndex        =   9
      Top             =   2040
      Width           =   4335
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PHONE NUMBER"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   6
      Top             =   8760
      Width           =   4935
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DATE OF BIRTH"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   5
      Top             =   7440
      Width           =   4935
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SEX"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   4
      Top             =   6120
      Width           =   4935
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   3
      Top             =   4800
      Width           =   4935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "USERNAME"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   2
      Top             =   3480
      Width           =   4935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NAME"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   1
      Top             =   2160
      Width           =   4935
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "STATE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10080
      TabIndex        =   7
      Top             =   7440
      Width           =   4935
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PROFESSION"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9720
      TabIndex        =   19
      Top             =   6240
      Width           =   4935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "INSTITUTE OF ENGINEERING AND MANAGEMENT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   23895
   End
   Begin VB.Image Image1 
      Height          =   11055
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   19935
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "COUNTRY"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   8
      Top             =   11280
      Width           =   4935
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub clear()
     Text1.Text = ""
    Text2.Text = ""
    DTPicker1.Value = "13/02/1997"
    Option1.Value = False
    Option2.Value = False
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    End Sub

Private Sub Command2_Click()
Dim conec As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    conec.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " & App.Path & "\Database.mdb; Persist Security Info=False"
    conec.CursorLocation = adUseClient
    conec.Open
    rs.Open "select * from users1", conec, adOpenDynamic, adLockPessimistic
    rs.AddNew
    rs.Fields("Names").Value = Text1.Text
    rs.Fields("Username").Value = Text2.Text
    rs.Fields("Password").Value = Text3.Text
    rs.Fields("Phone Number").Value = Val(Text4.Text)
    rs.Fields("Profession").Value = Combo1.List(Combo1.ListIndex)
    rs.Fields("State").Value = Text5.Text
    rs.Fields("Country").Value = Text6.Text
    rs.Fields("Date of Birth").Value = DTPicker1.Value
    If Option1.Value = True Then
    rs.Fields("Sex") = Option1.Caption
    Else
    rs.Fields("Sex") = Option2.Caption
    End If
    MsgBox "DATA RECORDED"
    rs.Update
    clear
    Unload Form1
    Form1.Show
    Form2.Hide
        End Sub

Private Sub Form_Load()
Image1.Picture = LoadPicture(App.Path & "\7.jpg")
clear
End Sub

