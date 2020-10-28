VERSION 5.00
Begin VB.Form Form02 
   Caption         =   "Admin Dashboard"
   ClientHeight    =   5535
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   7200
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
   ScaleHeight     =   5535
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "DataBase.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   5520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "StudentDetail"
      Top             =   4080
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate Result(For All)"
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   4680
      Width           =   3735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7095
   End
   Begin VB.Menu mnu_new 
      Caption         =   "New"
      Begin VB.Menu mnu_student 
         Caption         =   "Student"
      End
      Begin VB.Menu mnu_faculty 
         Caption         =   "Faculty"
      End
      Begin VB.Menu mnu_admin 
         Caption         =   "Admin"
      End
   End
   Begin VB.Menu mnu_edit 
      Caption         =   "Edit Database"
      Begin VB.Menu mnu_studedt 
         Caption         =   "Student Database"
      End
      Begin VB.Menu mnu_facedt 
         Caption         =   "Faculty Database"
      End
      Begin VB.Menu mnu_adedt 
         Caption         =   "Admin Database"
      End
   End
   Begin VB.Menu mnu_del 
      Caption         =   "Delete"
      Begin VB.Menu mnu_delstud 
         Caption         =   "All Student"
      End
      Begin VB.Menu mnu_delfac 
         Caption         =   "All Faculty"
      End
      Begin VB.Menu mnu_deladmn 
         Caption         =   "All Admins"
      End
      Begin VB.Menu mnu_reset 
         Caption         =   "Everything"
      End
   End
   Begin VB.Menu mnu_ep 
      Caption         =   "Edit Profile"
   End
   Begin VB.Menu mnu_exit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "Form02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Private Sub Command1_Click()
     Dim per, sum As Double
     Dim grad, str As String
     Data1.Refresh
     If Data1.Recordset.RecordCount = 0 Then
          Exit Sub
     End If
     Data1.Recordset.MoveFirst
     While (Not Data1.Recordset.EOF)
     sum = Data1.Recordset.Fields("Sub1").Value + Data1.Recordset.Fields("Sub2").Value + Data1.Recordset.Fields("Sub3").Value + Data1.Recordset.Fields("Sub4").Value + Data1.Recordset.Fields("Sub5").Value
     per = sum / 5
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
     str = "update studentdetail set percentage='" + CStr(per) + "', grade ='" + CStr(grad) + "' where id='" + Data1.Recordset.Fields("ID") + "';"
     Data1.Database.Execute (str)
     Data1.Recordset.MoveNext
     Wend
     Data1.Refresh
End Sub

Private Sub Form_Load()
     Label1.Caption = "Welcome " + Form01.Data1.Recordset.Fields("Name")
End Sub

Private Sub Form_Unload(Cancel As Integer)
     Unload Form02p1
     Unload Form02p2
     Unload Form02p3
     Unload Form02p5
     Unload Form02p4
     Unload Form01
     Load Form01
End Sub

Private Sub mnu_adedt_Click()
     Form02p7.Show
End Sub

Private Sub mnu_admin_Click()
     Form02p2.Show
End Sub

Private Sub mnu_deladmn_Click()
i = MsgBox("You will need to relogin. Continue?", vbYesNo)
If i = vbNo Then
     Exit Sub
End If
Data1.Database.Execute ("Delete * from AdminLogin;")
Data1.Database.Execute ("Insert into adminlogin values('1','1','Default Login');")
MsgBox ("All admins deleted")
Unload Me
End Sub

Private Sub mnu_delfac_Click()
i = MsgBox("You are going to delete all faculties. Continue?", vbYesNo)
If i = vbNo Then
     Exit Sub
End If
Data1.Database.Execute ("Delete * from FacultyLogin;")
MsgBox ("All faculties deleted")
End Sub

Private Sub mnu_delstud_Click()
i = MsgBox("You are going to delete all student records. Continue?", vbYesNo)
If i = vbNo Then
     Exit Sub
End If
Data1.Database.Execute ("Delete * from StudentLogin;")
Data1.Database.Execute ("Delete * from StudentDetail;")
MsgBox ("All students deleted")
End Sub

Private Sub mnu_ep_Click()
     Form02p4.Show
End Sub

Private Sub mnu_exit_Click()
     i = MsgBox("Do you want to exit?", vbYesNo)
     If i = vbYes Then
     Unload Me
     End If
End Sub

Private Sub mnu_facedt_Click()
     Form02p6.Show
End Sub

Private Sub mnu_faculty_Click()
     Form02p3.Show
End Sub

Private Sub mnu_reset_Click()
i = MsgBox("You are about to reset the system. Continue?", vbYesNo)
If i = vbNo Then
     Exit Sub
End If
Me.Enabled = False
Data1.Database.Execute ("Delete * from StudentLogin;")
Data1.Database.Execute ("Delete * from StudentDetail;")
Data1.Database.Execute ("Delete * from FacultyLogin;")
Data1.Database.Execute ("Delete * from AdminLogin;")
Data1.Database.Execute ("Insert into adminlogin values('1','1','Default Login');")
Me.Enabled = True
Unload Me
MsgBox ("System Reset completed")
End Sub

Private Sub mnu_studedt_Click()
 Form02p5.Show
End Sub

Private Sub mnu_student_Click()
     Form02p1.Show
End Sub
