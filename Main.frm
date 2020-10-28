VERSION 5.00
Begin VB.Form Form00 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Digital Examination System v0.0.0"
   ClientHeight    =   10425
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   16455
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
   MinButton       =   0   'False
   ScaleHeight     =   10425
   ScaleWidth      =   16455
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "Qp"
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   855
      Left            =   9120
      TabIndex        =   1
      Top             =   7560
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Log in"
      Height          =   855
      Left            =   4200
      TabIndex        =   0
      Top             =   7560
      Width           =   3375
   End
End
Attribute VB_Name = "Form00"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
     Form01.Show
End Sub

Private Sub Command2_Click()
     Unload Me
End Sub


Private Sub Command3_Click()
    Form03p1.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
     Unload Form01
     Unload Form02
     Unload Form02p1
     Unload Form02p2
     Unload Form02p3
     
End Sub
