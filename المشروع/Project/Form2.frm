VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«” „«—… «·„”«›—Ì‰"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11910
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form2.frx":0000
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
      Caption         =   "Õ–› »Ì«‰«  «·„”«›—Ì‰"
      Height          =   495
      Left            =   4568
      TabIndex        =   3
      Top             =   5160
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   " ÕœÌÀ »Ì«‰«  «·„”«›—Ì‰"
      Height          =   495
      Left            =   4568
      TabIndex        =   2
      Top             =   4440
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "«œŒ«· »Ì«‰«  «·„”«›—Ì‰"
      Height          =   495
      Left            =   4568
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   3720
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "⁄—÷ »Ì«‰«  «·„”«›—Ì‰"
      Height          =   495
      Left            =   4568
      TabIndex        =   0
      Top             =   3000
      Width           =   2775
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Mode = "Show"
Form2_0.Show 1
End Sub

Private Sub Command2_Click()
Form2.Hide
Form2_2.Show
End Sub

Private Sub Command3_Click()
Mode = "Edit"
Form2_0.Show 1

End Sub

Private Sub Command4_Click()
Mode = "Delete"
Form2_0.Show 1
End Sub

Private Sub Command5_Click()
Unload Form2
End Sub

Private Sub Form_Load()
Form1.Show
End Sub
