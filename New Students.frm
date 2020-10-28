VERSION 5.00
Begin VB.Form Form02p1 
   Caption         =   "New Student"
   ClientHeight    =   5325
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7920
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   5325
   ScaleWidth      =   7920
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "New Student"
      Height          =   5055
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      Begin VB.CommandButton Command2 
         Caption         =   "Back"
         Height          =   495
         Left            =   360
         TabIndex        =   12
         Top             =   4080
         Width           =   1575
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "DataBase.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   495
         Left            =   5040
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "StudentLogin"
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox Text5 
         Height          =   495
         Left            =   2640
         TabIndex        =   7
         Top             =   2760
         Width           =   2415
      End
      Begin VB.TextBox Text4 
         Height          =   465
         Left            =   2640
         TabIndex        =   6
         Top             =   2040
         Width           =   2295
      End
      Begin VB.TextBox Text3 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   1
         EndProperty
         Height          =   495
         Left            =   5760
         TabIndex        =   4
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   1
         EndProperty
         Height          =   495
         Left            =   3720
         TabIndex        =   3
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   480
         Left            =   1320
         TabIndex        =   2
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Create"
         Height          =   615
         Left            =   2640
         TabIndex        =   1
         Top             =   4080
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Password : (Confirm)"
         Height          =   855
         Left            =   240
         TabIndex        =   11
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Password :"
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "from :"
         Height          =   375
         Left            =   2400
         TabIndex        =   9
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Prefix :"
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "to"
         Height          =   375
         Left            =   5160
         TabIndex        =   5
         Top             =   840
         Width           =   255
      End
   End
End
Attribute VB_Name = "Form02p1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim from, tow As Integer
Dim fore, ident As String
Private Sub Command1_Click()
     fore = Text1.Text      'Prefix
     from = Val(Text2.Text) 'initial value
     If Text3.Text = Null Then
     tow = fore
     Else
     tow = Val(Text3.Text)  'final value
     End If
     If (allfill() <> True) Then
          Exit Sub
     End If
     Call newUser
End Sub

Function allfill() As Boolean
     allfill = False
     If Text2.Text = "" Or Text3.Text = "" Then
          MsgBox ("Enter all values")
     ElseIf Text4.Text = "" Then
          MsgBox ("Enter Password")
     ElseIf Text4.Text <> Text5.Text Then
          MsgBox ("Password do not match")
     ElseIf (from > tow) Then
          MsgBox ("Value Error")
     Else
          allfill = True
     End If
End Function

Sub newUser()
Dim i, count, fail As Integer
Dim str As String
count = 0
fail = 0
i = from
While (i <= tow)
    'check that id is not repeated and fill values
    str = "Insert into StudentLogin values('" + Text1.Text + CStr(i) + "', '" + Text5.Text + "', 'Student');"
    Data1.Database.Execute str
    str = "Insert into StudentDetail values('" + Text1.Text + CStr(i) + "', 'Student','0','0','0','0','0','0','0.0','N');"
    Data1.Database.Execute str
    i = i + 1
    count = count + 1
Wend
     Data1.Refresh
     clear
     MsgBox (CStr(count) + " New Students created" + CStr(fail) + "Entries failed")
End Sub


Private Sub Command2_Click()
     Unload Me
End Sub

Sub clear()
     Text1.Text = ""
     Text2.Text = ""
     Text3.Text = ""
     Text4.Text = ""
     Text5.Text = ""
End Sub
