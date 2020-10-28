VERSION 5.00
Begin VB.Form Form04 
   Caption         =   "Form6"
   ClientHeight    =   4740
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   8910
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form6"
   ScaleHeight     =   4740
   ScaleWidth      =   8910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Create new QP"
      Height          =   615
      Left            =   1080
      TabIndex        =   2
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Log Out"
      Height          =   375
      Left            =   7320
      TabIndex        =   1
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   7935
   End
   Begin VB.Menu flogout 
      Caption         =   "Log Out"
   End
End
Attribute VB_Name = "Form04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
     Call flogout_Click
End Sub

Private Sub Command2_Click()
     Form07.Show
End Sub

Private Sub flogout_Click()
     Form01.Show
     Unload Me
End Sub

Private Sub Form_Load()
     Label1.Caption = "Welcome " + Form01.Data3.Recordset.Fields("ID").Value
End Sub
