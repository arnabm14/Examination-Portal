VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Login"
   ClientHeight    =   9810
   ClientLeft      =   2415
   ClientTop       =   750
   ClientWidth     =   15375
   LinkTopic       =   "Form1"
   ScaleHeight     =   9810
   ScaleWidth      =   15375
   Begin VB.CommandButton Command1 
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1920
      TabIndex        =   6
      Top             =   7680
      Width           =   3975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SIGN UP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8640
      TabIndex        =   5
      Top             =   7800
      Width           =   3975
   End
   Begin VB.TextBox Text2 
      Height          =   855
      IMEMode         =   3  'DISABLE
      Left            =   8760
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   5400
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   8760
      TabIndex        =   1
      Top             =   3240
      Width           =   4335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "INSTITITE OF ENGINEERING AND MANAGEMENT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   11055
   End
   Begin VB.Label Label3 
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
      Height          =   855
      Left            =   1440
      TabIndex        =   2
      Top             =   5520
      Width           =   4335
   End
   Begin VB.Label Label2 
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
      Height          =   855
      Left            =   1560
      TabIndex        =   4
      Top             =   3240
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   9735
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Command1_Click()
Dim strName As String
Dim strPass As String
Dim a As Integer
strName = Text1.Text
strPass = Text2.Text
Dim conec As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim b
conec.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " & App.Path & "\Database.mdb; Persist Security Info=False"
conec.CursorLocation = adUseClient
conec.Open
rs.Open "select Username, Password, Profession from users1", conec, adOpenDynamic, adLockPessimistic, adcmdtxt
Do Until rs.EOF
If rs.Fields("Username").Value = strName And rs.Fields("Password").Value = strPass Then
        a = 1
        If rs.Fields("Profession").Value = "STUDENT" Then
        b = 1
        Else
        b = 2
        End If
    Exit Do
    Else
    rs.MoveNext
    a = 2
    End If
Loop
If a = 2 Then
MsgBox "Invalid User Name or Password", vbOKOnly
ElseIf a = 1 Then
    If b = 1 Then
    Form1.Hide
    dForm1.Show
    Else
    Form1.Hide
    aForm1.Show
    End If
End If
End Sub

Private Sub Command2_Click()
Form1.Hide
Form2.Show
End Sub

Private Sub Form_Load()
Image1.Picture = LoadPicture(App.Path & "\3.jpg")
End Sub
