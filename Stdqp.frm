VERSION 5.00
Begin VB.Form dForm2 
   Appearance      =   0  'Flat
   Caption         =   "Exam"
   ClientHeight    =   9045
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   15240
   LinkTopic       =   "Form2"
   ScaleHeight     =   9045
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Next"
      Height          =   495
      Left            =   11160
      TabIndex        =   12
      Top             =   6600
      Width           =   3015
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      DataField       =   "Fourth_option"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3600
      TabIndex        =   11
      Text            =   "Text5"
      Top             =   6600
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      DataField       =   "Third_option"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3600
      TabIndex        =   10
      Text            =   "Text4"
      Top             =   5640
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      DataField       =   "Second_option"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   3600
      TabIndex        =   9
      Text            =   "Text3"
      Top             =   4800
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      DataField       =   "First_option"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   3840
      Width           =   1335
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H80000014&
      Caption         =   "Option4"
      Height          =   615
      Left            =   3120
      TabIndex        =   6
      Top             =   6480
      Width           =   255
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H80000014&
      Caption         =   "Option3"
      Height          =   495
      Left            =   3120
      TabIndex        =   5
      Top             =   5520
      Width           =   255
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H80000014&
      Caption         =   "Option2"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   4680
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H80000014&
      Caption         =   "Option1"
      Height          =   615
      Left            =   3120
      TabIndex        =   3
      Top             =   3600
      Width           =   135
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      DataField       =   "Question"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   960
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1320
      Width           =   13335
   End
   Begin VB.Timer Timer2 
      Interval        =   60000
      Left            =   14040
      Top             =   1920
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Submit"
      Height          =   615
      Left            =   6480
      TabIndex        =   1
      Top             =   8280
      Width           =   2895
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   14280
      Top             =   3960
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      Height          =   495
      Left            =   12960
      TabIndex        =   0
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Time Left :="
      Height          =   375
      Left            =   11760
      TabIndex        =   7
      Top             =   480
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   9015
      Left            =   0
      Top             =   0
      Width           =   15255
   End
End
Attribute VB_Name = "dForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, b, c, d
Dim num_question
Dim l, ans
Dim score






Private Sub Command1_Click()
l = 0

dForm3.Label7.Caption = score
Unload dForm2
dForm2.Hide
Unload Form1
Form1.Show
dForm3.Show
Timer1.Enabled = False
Timer2.Enabled = False
End Sub



Private Sub Command2_Click()

Dim connec2 As New ADODB.Connection
Dim rs2 As New ADODB.Recordset
connec2.ConnectionString = "provider=Microsoft.jet.OLEDB.4.0;Data Source = " & App.Path & "\Database.mdb;Persist Security Info=False"
connec2.CursorLocation = adUseClient
connec2.Open
rs2.Open "select Question,First_option,Second_option,Third_option,Fourth_option,Correct_option,Answer from table1", connec2, adOpenDynamic, adLockBatchOptimistic, adcmdtxt
Dim s
s = 0
   
Do Until s = d - 1
    If Not rs2.EOF Then
    rs2.MoveNext
    s = s + 1
    Else
    Exit Do
    End If
Loop






If rs2.Fields("Correct_option").Value = ans Then
    score = score + 1
End If

rs2.MoveNext



If Not rs2.EOF Then
Text1.Text = rs2.Fields("Question").Value
Text2.Text = rs2.Fields("First_option").Value
Text3.Text = rs2.Fields("Second_option").Value
Text4.Text = rs2.Fields("Third_option").Value
Text5.Text = rs2.Fields("Fourth_option").Value
d = d + 1
End If


End Sub

Private Sub Form_Load()
Image1.Picture = LoadPicture(App.Path & "\5.jpg")
d = 1
ans = 0
score = 0
Dim connec As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim a
connec.ConnectionString = "provider=Microsoft.jet.OLEDB.4.0;Data Source = " & App.Path & "\Database.mdb;Persist Security Info=False"
connec.CursorLocation = adUseClient
connec.Open
rs.Open "select Question from Table1", connec, adOpenDynamic, adLockBatchOptimistic, adcmdtxt
Do Until rs.EOF
    num_question = num_question + 1
    rs.MoveNext
Loop
Dim connec1 As New ADODB.Connection
Dim rs1 As New ADODB.Recordset

connec1.ConnectionString = "provider=Microsoft.jet.OLEDB.4.0;Data Source = " & App.Path & "\Database.mdb;Persist Security Info=False"
connec1.CursorLocation = adUseClient
connec1.Open
rs1.Open "select Question,First_option,Second_option,Third_option,Fourth_option from table1", connec1, adOpenDynamic, adLockBatchOptimistic, adcmdtxt
Text1.Text = rs1.Fields("Question").Value
Text2.Text = rs1.Fields("First_option").Value
Text3.Text = rs1.Fields("Second_option").Value
Text4.Text = rs1.Fields("Third_option").Value
Text5.Text = rs1.Fields("Fourth_option").Value
rs1.Close

End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
ans = 1
End If
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
ans = 2
End If
End Sub

Private Sub Option3_Click()
If Option3.Value = True Then
ans = 3
End If
End Sub

Private Sub Option4_Click()
If Option4.Value = True Then
ans = 4
End If
End Sub

Private Sub Timer1_Timer()
Label1.Caption = Label1.Caption - 1
If Label1.Caption = 0 Then
dForm2.Hide
dForm3.Label7.Caption = score
Unload Form1
Form1.Show
dForm3.Show
End If
End Sub

Private Sub Timer2_Timer()
dForm3.Label7.Caption = score
dForm3.Show
Timer2.Enabled = False
End Sub
