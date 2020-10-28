VERSION 5.00
Begin VB.Form Form03 
   Caption         =   "Form4"
   ClientHeight    =   6720
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12150
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   ScaleHeight     =   6720
   ScaleWidth      =   12150
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Subject Details :"
      Height          =   3735
      Left            =   6240
      TabIndex        =   13
      Top             =   1920
      Width           =   5295
   End
   Begin VB.Label Label5 
      Height          =   375
      Left            =   6480
      TabIndex        =   14
      Top             =   1440
      Width           =   4935
   End
   Begin VB.Label Label3 
      Height          =   375
      Index           =   5
      Left            =   2400
      TabIndex        =   12
      Top             =   4560
      Width           =   3495
   End
   Begin VB.Label Label3 
      Height          =   375
      Index           =   4
      Left            =   2400
      TabIndex        =   11
      Top             =   4080
      Width           =   3495
   End
   Begin VB.Label Label3 
      Height          =   375
      Index           =   3
      Left            =   2400
      TabIndex        =   10
      Top             =   3480
      Width           =   3495
   End
   Begin VB.Label Label3 
      Height          =   375
      Index           =   2
      Left            =   2400
      TabIndex        =   9
      Top             =   2760
      Width           =   3495
   End
   Begin VB.Label Label3 
      Height          =   375
      Index           =   1
      Left            =   2400
      TabIndex        =   8
      Top             =   2160
      Width           =   3495
   End
   Begin VB.Label Label3 
      Height          =   375
      Index           =   0
      Left            =   2520
      TabIndex        =   7
      Top             =   1560
      Width           =   3495
   End
   Begin VB.Label Label2 
      Caption         =   "Division :"
      Height          =   375
      Index           =   5
      Left            =   0
      TabIndex        =   6
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Class :"
      Height          =   375
      Index           =   4
      Left            =   0
      TabIndex        =   5
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Course :"
      Height          =   375
      Index           =   3
      Left            =   0
      TabIndex        =   4
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Contact :"
      Height          =   375
      Index           =   2
      Left            =   0
      TabIndex        =   3
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Address :"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Full Name :"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
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
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   11895
   End
End
Attribute VB_Name = "Form03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim i As Integer
Dim sem As String

Private Sub Command1_Click()
     Form03.Visible = True
     Me.Visible = False
End Sub

Private Sub Form_Load()
     If Format(Now, "MM") < 6 Then
     sem = " Semester I"
     Else
     sem = " Semester II"
     End If
     Label1.Caption = "Welcome " + Form01.Data2.Recordset.Fields(0)
     i = 0
     Do While (i < 6)
          Label3(i).Caption = Form01.Data2.Recordset.Fields(i).Value
          i = i + 1
     Loop
     Label5.Caption = Form01.Data2.Recordset.Fields("class") + sem
End Sub

