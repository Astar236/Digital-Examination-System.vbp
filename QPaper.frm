VERSION 5.00
Begin VB.Form Form03p1 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Form3"
   ClientHeight    =   7125
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9510
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   ScaleHeight     =   7125
   ScaleWidth      =   9510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Submit"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7080
      TabIndex        =   25
      Top             =   4200
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   495
      Index           =   0
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1
      Top             =   360
   End
   Begin VB.CommandButton Command3 
      Caption         =   "<<Previous"
      Enabled         =   0   'False
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Next>>"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5640
      TabIndex        =   3
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Begin"
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "DataBase.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   7320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "QSub1"
      Top             =   6120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Height          =   3375
      Left            =   600
      TabIndex        =   0
      Top             =   960
      Width           =   6135
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   495
         Index           =   4
         Left            =   4680
         TabIndex        =   8
         Top             =   2400
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   495
         Index           =   3
         Left            =   3240
         TabIndex        =   7
         Top             =   2400
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   495
         Index           =   2
         Left            =   1560
         TabIndex        =   6
         Top             =   2400
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   495
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   4935
      End
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Time Remaining = 10 : 00"
      Height          =   375
      Left            =   5280
      TabIndex        =   23
      Top             =   5520
      Width           =   3375
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Skipped"
      Height          =   375
      Left            =   5880
      TabIndex        =   22
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "Answered"
      Height          =   375
      Left            =   3200
      TabIndex        =   21
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "Unanswered"
      Height          =   375
      Left            =   480
      TabIndex        =   20
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "!!!All The Best!!!"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   19
      Top             =   240
      Width           =   8175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "10"
      Enabled         =   0   'False
      Height          =   375
      Index           =   9
      Left            =   8160
      TabIndex        =   18
      Top             =   3660
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "9"
      Enabled         =   0   'False
      Height          =   375
      Index           =   8
      Left            =   7080
      TabIndex        =   17
      Top             =   3660
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "8"
      Enabled         =   0   'False
      Height          =   375
      Index           =   7
      Left            =   8160
      TabIndex        =   16
      Top             =   3045
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "7"
      Enabled         =   0   'False
      Height          =   375
      Index           =   6
      Left            =   7080
      TabIndex        =   15
      Top             =   3045
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "6"
      Enabled         =   0   'False
      Height          =   375
      Index           =   5
      Left            =   8160
      TabIndex        =   14
      Top             =   2430
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "5"
      Enabled         =   0   'False
      Height          =   375
      Index           =   4
      Left            =   7080
      TabIndex        =   13
      Top             =   2430
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "4"
      Enabled         =   0   'False
      Height          =   375
      Index           =   3
      Left            =   8160
      TabIndex        =   12
      Top             =   1815
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "3"
      Enabled         =   0   'False
      Height          =   375
      Index           =   2
      Left            =   7080
      TabIndex        =   11
      Top             =   1815
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "2"
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   8160
      TabIndex        =   10
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "1"
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   7080
      TabIndex        =   9
      Top             =   1200
      Width           =   615
   End
End
Attribute VB_Name = "Form03p1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim ans(1 To 11), mks(1 To 11), total, i, min, sec As Integer


Private Sub Command1_Click()  'Initiates Recordset and calls newQs to load options
     If Data1.Recordset.RecordCount > 10 Then
     MsgBox ("This Question Paper is Not Yet Fully Created" + CStr(Data1.Recordset.RecordCount))
     Exit Sub
     End If
     Command1.Enabled = False
     Timer1.Enabled = True
     Command1.Visible = False
     i = 1
     Do While (i < 5)
          Option1(i).Enabled = True
          i = i + 1
     Loop
     i = 0
     Do While (i < 10)
          Label2(i).Enabled = True
          i = i + 1
     Loop
     Command2.Enabled = True
     Command3.Enabled = True
     Command4.Enabled = True
     
     Data1.Recordset.MoveFirst
     Call NewQs
     
End Sub

Private Sub Command2_Click() 'move to next record and call newqs
     If ans(Val(Frame1.Caption)) = -1 Then
     ans(Val(Frame1.Caption)) = 0
     End If
     Data1.Recordset.MoveNext
     If Data1.Recordset.EOF Then
     Call ColorCode
          i = MsgBox("You Have attempted all the questions. Submit Response?", vbYesNo + vbInformation, "Super")
          Data1.Recordset.MoveLast
          If i = vbYes Then
               Call SubmitResponse 'update db
               Unload Me
          End If
     Else
          Call NewQs
     End If
End Sub

Private Sub Command3_Click()  'move to prev record and change text on form
     If ans(Val(Frame1.Caption)) = -1 Then
     ans(Val(Frame1.Caption)) = 0
     End If
     Data1.Recordset.MovePrevious
     If Data1.Recordset.BOF Then
          Data1.Recordset.MoveFirst
     Else
          Call NewQs
     End If
End Sub

Function NewQs()
     Call ColorCode 'Upadate colorcode as per answered q's
     Frame1.Caption = Data1.Recordset.Fields("Srno")
     Label1.Caption = Data1.Recordset.Fields("Question")
     i = 1
     Do While (i < 5)
          Option1(i).Caption = Data1.Recordset.Fields(1 + i)
          Option1(i).Value = False
          i = i + 1
     Loop           'Q's and optn on form updated
     If ans(Val(Frame1.Caption)) > 0 Then   '
          Option1(ans(Val(Frame1.Caption)) - 1).Value = True
     End If
End Function

Private Sub Command4_Click()
     Call SubmitResponse
End Sub

Private Sub Form_Load()
     i = 1
     Do While (i < 11)
          ans(i) = -1
          i = i + 1
     Loop
     sec = 0
min = 10
End Sub

Private Sub ColorCode()  'Decides color of label2
     i = 1
     While (i < 11)
          If ans(i) = -1 Then
               Label2(i - 1).BackColor = Label4.BackColor
          ElseIf ans(i) > 0 Then
               Label2(i - 1).BackColor = Label5.BackColor
          Else
               Label2(i - 1).BackColor = Label6.BackColor
          End If
          i = i + 1
     Wend
End Sub


Private Sub Form_Unload(Cancel As Integer)
     MsgBox ("Response not Submitted")
     'Form04.Visible = True
End Sub

Private Sub Label2_Click(Index As Integer)
Dim j As Integer
     j = Val(Frame1.Caption)
     Data1.Recordset.MoveFirst
     i = 0
     Do While (i < Index)
          Data1.Recordset.MoveNext
          If Data1.Recordset.EOF Then
               MsgBox ("Record not found")
               Call Label2_Click(j - 1)
               i = Index
          End If
          i = i + 1
     Loop
     NewQs
End Sub


Private Sub Option1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

     If Option1(Index).Value = True Then
          ans(Val(Frame1.Caption)) = 0
          Option1(Index).Value = False
     Else
          Option1(Index).Value = True
          ans(Val(Frame1.Caption)) = Index + 1
     End If
     If Option1(Index).Caption = Data1.Recordset.Fields("Answer").Value Then
          mks(Frame1.Caption) = 4
     Else
          mks(Frame1.Caption) = 0
     End If
End Sub

Private Sub Timer1_Timer()    'Updates the timer and Unloads the form after 10 mins
     
     If sec = 0 Then
          sec = 60
          min = min - 1
     End If
     sec = sec - 1
     Label7.Caption = "Time Remaining = " & min & " : " & sec
     If min = 0 And sec = 0 Then
          MsgBox ("You Have Exceeded the time Limit")
          Call SubmitResponse
     End If
End Sub

Private Function SubmitResponse()  'Calculate total from mks array and update in db

     i = 1
     total = 0
     Do While (i < 11)
     total = mks(i) + total
     i = i + 1
     Loop
     i = MsgBox("Response Submitted Successfully. Total Score= " & total)
     'update db
     'Form04.Visible = True
     Unload Me
End Function
