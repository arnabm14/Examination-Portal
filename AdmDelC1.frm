VERSION 5.00
Begin VB.Form aForm5 
   Caption         =   "Admin"
   ClientHeight    =   5610
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9795
   LinkTopic       =   "Form5"
   ScaleHeight     =   5610
   ScaleWidth      =   9795
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "No"
      Height          =   495
      Left            =   6840
      TabIndex        =   2
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Yes"
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Are You sure you want to delete this Question?"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   9735
   End
   Begin VB.Image Image1 
      Height          =   5655
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000014&
      Height          =   2295
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   8175
   End
End
Attribute VB_Name = "aForm5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Str As Integer
Dim conec As New ADODB.Connection
Dim rs As New ADODB.Recordset
    
Private Sub Command1_Click()
rs.Close
rs.Open "delete from Table1 where Qno=" & Str & "", conec, adOpenDynamic, adLockPessimistic
conec.Close
aForm5.Hide
aForm1.Refresh
aForm1.Show
End Sub

Private Sub Command2_Click()
aForm5.Hide
aForm1.Refresh
aForm1.Show

End Sub

Private Sub Form_Load()
Image1.Picture = LoadPicture(App.Path & "\4.jpg")
Str = Val(aForm4.Text1.Text)

    conec.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " & App.Path & "\Database.mdb; Persist Security Info=False"
    conec.CursorLocation = adUseClient
    conec.Open
    rs.Open "select * from Table1", conec, adOpenDynamic, adLockPessimistic
    Dim a As Integer
    a = 0
    Do Until rs.EOF
        If rs.Fields("Qno").Value = Str Then
            Label1.Caption = rs.Fields("Question").Value
            a = 1
            Exit Do
        Else
            rs.MoveNext
        End If
    Loop
    If a = "0" Then
    rs.Close
    conec.Close
    aForm5.Hide
    aForm1.Refresh
    aForm1.Show
    Label1.Caption = "Press No to continue"
    MsgBox "Enter a valid Question number"
    End If
End Sub

