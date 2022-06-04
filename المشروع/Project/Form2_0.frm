VERSION 5.00
Begin VB.Form Form2_0 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " ÕœÌœ «·‘Œ’ «·„ÿ·Ê»"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3930
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   1650
   ScaleWidth      =   3930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "«·€«¡ «·«„—"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "„Ê«›ﬁ"
      Default         =   -1  'True
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "«œŒ· —„“ «·‘Œ’"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "Form2_0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Mode = "Show" And Text1.Text <> "" Then
    StrSearch = Text1.Text
    Unload Form2_0
    Form2_1.Show
ElseIf Mode = "Edit" And Text1.Text <> "" Then
    StrSearch = Text1.Text
    Unload Form2_0
    Form2_3.Show
ElseIf Mode = "Delete" And Text1.Text <> "" Then
    StrSearch = Text1.Text
    Unload Form2_0
    Form2_4.Show
End If
End Sub

Private Sub Command2_Click()
Unload Form2_0
End Sub
