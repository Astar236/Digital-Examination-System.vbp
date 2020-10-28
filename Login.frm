VERSION 5.00
Begin VB.Form Form01 
   ClientHeight    =   7890
   ClientLeft      =   6930
   ClientTop       =   2205
   ClientWidth     =   6615
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
   MaxButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   6615
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   "DataBase.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7320
      Width           =   1140
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   615
      Left            =   480
      TabIndex        =   4
      Top             =   7080
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Login Frame"
      Height          =   5415
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   5895
      Begin VB.Data Data3 
         Caption         =   "Data3"
         Connect         =   "Access"
         DatabaseName    =   "DataBase.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   420
         Left            =   4200
         Options         =   0
         ReadOnly        =   -1  'True
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "FacultyLogin"
         Top             =   4680
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data2 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   "DataBase.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   420
         Left            =   4200
         Options         =   0
         ReadOnly        =   -1  'True
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "StudentLogin"
         Top             =   4080
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Admin"
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   2400
         PasswordChar    =   "*"
         TabIndex        =   8
         ToolTipText     =   "Your Password"
         Top             =   2400
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   2400
         TabIndex        =   7
         ToolTipText     =   "The ID you receieved from your Admin"
         Top             =   1440
         Width           =   2535
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Faculty"
         Height          =   495
         Left            =   3240
         TabIndex        =   6
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Student"
         Height          =   495
         Left            =   1800
         TabIndex        =   5
         Top             =   480
         Width           =   1455
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "DataBase.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   420
         Left            =   4200
         Options         =   0
         ReadOnly        =   -1  'True
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "AdminLogin"
         Top             =   3600
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Login"
         Height          =   735
         Left            =   1680
         TabIndex        =   3
         Top             =   4440
         Width           =   2295
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Show Password"
         Height          =   315
         Left            =   1080
         TabIndex        =   2
         Top             =   3240
         Width           =   3255
      End
      Begin VB.Label Label2 
         Caption         =   "User ID :"
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Password :"
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   2400
         Width           =   2055
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Login Page"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   6015
   End
End
Attribute VB_Name = "Form01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()    'Sets Password Character of Text Box 2
     If Check1.Value = vbChecked Then
          Text2.PasswordChar = ""
     Else
          Text2.PasswordChar = "*"
     End If
End Sub

Private Sub Command1_Click()
     If Text1.Text = "" Then  'Checks if Id & Password field is filled or not else Exits sub
          i = MsgBox("Enter Username", vbOKOnly + vbExclamation, "Error")
          Exit Sub
     End If
     If Text2.Text = "" Then
          i = MsgBox("Enter Password", vbOKOnly + vbExclamation, "Error")
          Exit Sub
     End If         'Check compllete
     
     If Option1.Value = True Then  'Validates ID & Password & opens appropriate form
          resp = search(Data1)
          If resp = True Then
               i = MsgBox("Welcome Admin", vbOKOnly + vbInformation, "Success!!!")
               Form02.Show
               Form01.Hide
               clear
          End If
     ElseIf Option2.Value = True Then
          resp = search(Data2)
          If resp = True Then
               i = MsgBox("Welcome Student", vbOKOnly + vbInformation, "Success!!!")
               Form03.Show
               Form01.Hide
               clear
          End If
     Else
          resp = search(Data3)
          If resp = True Then
               i = MsgBox("Welcome Faculty", vbOKOnly + vbInformation, "Success!!!")
               Form04.Show
               Form01.Hide
               clear
          End If
     End If                        'Validation Complete

End Sub
Function clear() As Integer   'Clears textboxes
Text1.Text = ""
Text2.Text = ""
End Function

Private Sub Command2_Click()  'Exit Command
     Unload Me
End Sub


Private Sub Form_Load()
     Call clear
End Sub

Function search(db As Data) As Boolean  'Searches for a record and returns true if found it
     search = False
     db.Recordset.MoveFirst
     If db.Recordset.EOF Then
          MsgBox ("Data ERROR")
     End If
     Do Until db.Recordset.EOF
          If db.Recordset.Fields("ID").Value = Text1.Text Then
               If db.Recordset.Fields("Password").Value = Text2.Text Then
                    search = True
                    Exit Function
               Else
                    i = MsgBox("Password is Incorrect", vbRetryCancel + vbCritical, "Error")
                    Exit Function
               End If
          End If
          db.Recordset.MoveNext
     Loop
     
     i = MsgBox("Username is Incorrect", vbRetryCancel + vbCritical, "Error")
     If i = vbCancel Then
          Call clear
     End If
     
End Function

