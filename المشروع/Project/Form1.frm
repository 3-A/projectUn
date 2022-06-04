VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "‰Ÿ«„ ‘—ﬂ… ”Ì«ÕÌ…"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Œ—ÊÃ"
      Height          =   495
      Left            =   4568
      TabIndex        =   4
      Top             =   5880
      Width           =   2775
   End
   Begin VB.CommandButton Command4 
      Caption         =   "«⁄œ«œ "
      Height          =   495
      Left            =   4568
      TabIndex        =   3
      Top             =   5160
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "«” „«—… «·”›—« "
      Height          =   495
      Left            =   4568
      TabIndex        =   2
      Top             =   4440
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "«” „«—… «·›‰«œﬁ"
      Height          =   495
      Left            =   4568
      TabIndex        =   1
      Top             =   3720
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "«” „«—… «·„”«›—Ì‰"
      Height          =   495
      Left            =   4568
      TabIndex        =   0
      Top             =   3000
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Hide
Form2.Show
End Sub

Private Sub Command2_Click()
Form1.Hide
Form3.Show

End Sub

Private Sub Command3_Click()
Form1.Hide
Form4.Show

End Sub

Private Sub Command4_Click()
Form1.Hide
Form0.Show
End Sub

Private Sub Command5_Click()
End
End Sub
