VERSION 5.00
Begin VB.Form Form3_0 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " ÕœÌœ «·›‰œﬁ «·„ÿ·Ê»"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3795
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   1680
   ScaleWidth      =   3795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "„Ê«›ﬁ"
      Default         =   -1  'True
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "«·€«¡ «·«„—"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "«œŒ· —„“ «·›‰œﬁ"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "Form3_0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Mode = "Show" And Text1.Text <> "" Then
    StrSearch = Text1.Text
    Unload Form3_0
    Form3_1.Show
ElseIf Mode = "Delete" And Text1.Text <> "" Then
    StrSearch = Text1.Text
    Unload Form3_0
    Form3_2.Show
End If
End Sub

Private Sub Command2_Click()
Unload Form3_0
End Sub

