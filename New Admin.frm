VERSION 5.00
Begin VB.Form Form02p2 
   Caption         =   "New Admin"
   ClientHeight    =   5445
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6765
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   6765
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "New Admin Details"
      Height          =   4575
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   6135
      Begin VB.CommandButton Command2 
         Caption         =   "Back"
         Height          =   315
         Left            =   3480
         TabIndex        =   10
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "DataBase.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   420
         Left            =   3840
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "AdminLogin"
         Top             =   3240
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Text4 
         Height          =   435
         Left            =   1680
         TabIndex        =   9
         Top             =   2640
         Width           =   3015
      End
      Begin VB.TextBox Text3 
         Height          =   435
         Left            =   1680
         TabIndex        =   8
         Top             =   1920
         Width           =   3015
      End
      Begin VB.TextBox Text2 
         Height          =   435
         Left            =   1680
         TabIndex        =   7
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Height          =   435
         Left            =   1680
         TabIndex        =   6
         Top             =   600
         Width           =   3015
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Create"
         Height          =   495
         Left            =   840
         TabIndex        =   5
         Top             =   3720
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "Password : (Confirm)"
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Password :"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "ID :"
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Name :"
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   2055
      End
   End
End
Attribute VB_Name = "Form02p2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
     Dim str As String
     If (allfill() = False) Then
          Exit Sub
     ElseIf (preexist()) Then
          MsgBox ("Id Exists. use new one")
          Exit Sub
     Else
     str = "Insert into AdminLogin values('" + Text2.Text + "', '" + Text3.Text + "', '" + Text1.Text + "');"
     Data1.Database.Execute str
     Data1.Refresh
     clear
     MsgBox ("Admin Created")
     End If
End Sub

 Function allfill() As Boolean     'checks if all fields are filled or not
     allfill = False
     If Text1.Text = "" Then
     MsgBox ("Enter Name")
     ElseIf Text2.Text = "" Then
     MsgBox ("Enter ID")
     ElseIf Text3.Text = "" Then
     MsgBox ("Enter Password")
     ElseIf Text4.Text = "" Then
     MsgBox ("Reenter password")
     ElseIf Text3.Text <> Text4.Text Then
     MsgBox ("Passwords do not match")
     Else
     allfill = True
     End If
End Function

Function preexist() As Boolean     'checks if id already exists or not
     preexist = False
     
     If Data1.Recordset.RecordCount = 0 Then
     preexist = False
     Else
          Data1.Recordset.MoveFirst
          Do Until (Data1.Recordset.EOF)
          If Data1.Recordset.Fields("ID") = Text2.Text Then
               preexist = True
          End If
          Data1.Recordset.MoveNext
          Loop
     End If
End Function

Sub clear()
     Text1.Text = ""
     Text2.Text = ""
     Text3.Text = ""
     Text4.Text = ""
End Sub

Private Sub Command2_Click()
     Unload Me
End Sub
