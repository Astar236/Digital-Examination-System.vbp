VERSION 5.00
Begin VB.Form Form02p4 
   Caption         =   "Edit Profile"
   ClientHeight    =   5490
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8280
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
   ScaleHeight     =   5490
   ScaleWidth      =   8280
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "DataBase.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   435
      Left            =   2880
      TabIndex        =   7
      Top             =   2640
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   435
      Left            =   2880
      TabIndex        =   6
      Top             =   1920
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit without Saving"
      Height          =   615
      Left            =   1200
      TabIndex        =   2
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save and Exit"
      Height          =   495
      Left            =   5520
      TabIndex        =   1
      Top             =   4680
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "New Password"
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "New Name"
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1320
      Width           =   6495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   6495
   End
End
Attribute VB_Name = "Form02p4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim edited As Boolean

Private Sub Command1_Click()
     Dim str As String
     If Text1.Text = "" Or Text2.Text = "" Then
          MsgBox ("Enter new field values to update")
          Exit Sub
     End If
     i = MsgBox("You will need to re login.Continue?", vbYesNo)
     If i = vbNo Then
          Exit Sub
     End If
     edited = True
     str = "update AdminLogin set Name='" + CStr(Text1.Text) + "', Password='" + CStr(Text2.Text) + "' where ID='" + Form01.Data1.Recordset.Fields("ID") + "';"
     Data1.Database.Execute (str)
     Unload Me
End Sub

Private Sub Command2_Click()
     Unload Me
End Sub

Private Sub Form_Load()
     edited = False
     Label1.Caption = "Admin : " + Form01.Data1.Recordset.Fields("Name")
     Label2.Caption = "ID : " + Form01.Data1.Recordset.Fields("ID")
     Text1.Text = Form01.Data1.Recordset.Fields("Name")
     Text2.Text = Form01.Data1.Recordset.Fields("Password")
     
End Sub

Private Sub Form_Unload(Cancel As Integer)
     If edited = True Then
     Unload Form02
     Unload Form02p1
     Unload Form02p2
     Unload Form02p3
     Form01.Show
     End If
End Sub

