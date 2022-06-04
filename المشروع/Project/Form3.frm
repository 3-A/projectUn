VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«” „«—… «·›‰«œﬁ"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11910
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form3.frx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Œ—ÊÃ"
      Height          =   495
      Left            =   4568
      TabIndex        =   2
      Top             =   5100
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Õ–› »Ì«‰«  «·›‰«œﬁ"
      Height          =   495
      Left            =   4568
      TabIndex        =   1
      Top             =   4380
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "⁄—÷ »Ì«‰«  «·›‰«œﬁ"
      Height          =   495
      Left            =   4568
      TabIndex        =   0
      Top             =   3660
      Width           =   2775
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Mode = "Show"
Form3_0.Show 1
End Sub

Private Sub Command2_Click()
Mode = "Delete"
Form3_0.Show 1

End Sub

Private Sub Command3_Click()
Unload Form3
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.Show
End Sub
