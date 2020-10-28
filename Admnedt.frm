VERSION 5.00
Begin VB.Form Form02p7 
   Caption         =   "Edit Admin Records"
   ClientHeight    =   8400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9930
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   9930
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Back"
      Height          =   495
      Left            =   6240
      TabIndex        =   7
      Top             =   3240
      Width           =   2895
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete Record"
      Height          =   495
      Left            =   2520
      TabIndex        =   6
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Data Data1 
      Caption         =   "Control"
      Connect         =   "Access"
      DatabaseName    =   "DataBase.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   465
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "AdminLogin"
      Top             =   3240
      Width           =   3660
   End
   Begin VB.TextBox Text2 
      DataField       =   "Password"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   2400
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      DataField       =   "Name"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   1800
      Width           =   4695
   End
   Begin VB.Label Label14 
      DataField       =   "ID"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   840
      TabIndex        =   8
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Password :"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Name :"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "ID :"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Table Name: Admin Login"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9615
   End
End
Attribute VB_Name = "Form02p7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
     Dim s1, s2, s3, s4, s5 As Integer
     Dim per As Double
     Dim grad As String
     s1 = Val(Text3.Text)
     s2 = Val(Text4.Text)
     s3 = Val(Text5.Text)
     s4 = Val(Text6.Text)
     s5 = Val(Text7.Text)
     per = (s1 + s2 + s3 + s4 + s5) / 5
     Label12.Caption = per
     Select Case per
     
     Case Is > 90:
     grad = "A+"
     
     Case Is > 80:
     grad = "A"
     
     Case Is > 70:
     grad = "B+"
     
     Case Is > 60:
     grad = "B"
     
     Case Is > 50:
     grad = "C"
     Case Is > 40:
     grad = "D"
     Case Is > 34
     grad = "E"
     Case Is < 35
     grad = "F"
     End Select
     Label13.Caption = grad
End Sub

Private Sub Command3_Click()

     Dim str, id As String
     On Error GoTo error
     id = Data1.Recordset.Fields("ID")
     If Data1.Recordset.RecordCount = 1 Then
     MsgBox ("Must have at least one admin")
     Exit Sub
     ElseIf id = Form01.Data1.Recordset.Fields("ID") Then
     MsgBox ("You cant delete yourself")
     Exit Sub
     End If
     
     str = "DELETE FROM AdminLogin WHERE ID = '" + id + "';"
     Data1.Database.Execute (str)
     Data1.Refresh
     MsgBox ("Record deleted sucessfully")
error:
End Sub

Private Sub Command4_Click()
     Unload Me
End Sub
