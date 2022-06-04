VERSION 5.00
Begin VB.Form Form4_1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "⁄—÷ »Ì«‰«  «·”›—« "
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11910
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Align           =   2  'Align Bottom
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\WINDOWS\Company.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Journey"
      RightToLeft     =   -1  'True
      Top             =   8175
      Width           =   11910
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Œ—ÊÃ"
      Height          =   495
      Left            =   5348
      TabIndex        =   4
      Top             =   7320
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      DataField       =   "Size"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   360
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   2880
      Width           =   9375
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      DataField       =   "Mony"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   360
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   2040
      Width           =   9375
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      DataField       =   "Date"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   360
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1200
      Width           =   9375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "Id"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   360
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   9375
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "„œ… «·”›—…:"
      Height          =   375
      Index           =   3
      Left            =   10080
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "”⁄— »ÿ«ﬁ… «·”›—…:"
      Height          =   375
      Index           =   2
      Left            =   10080
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   " «—ÌŒ «·”›—…:"
      Height          =   375
      Index           =   1
      Left            =   10080
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "—„“ «·”›—…:"
      Height          =   375
      Index           =   0
      Left            =   10080
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "Form4_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Form4_1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form4.Show
End Sub
Private Sub Data1_Reposition()
Data1.Caption = "«·”Ã· " & CStr(Data1.Recordset.AbsolutePosition + 1) & " „‰ " & CStr(Data1.Recordset.RecordCount)
End Sub
Private Sub Form_Activate()
Data1.Recordset.MoveLast
Data1.Recordset.MoveFirst
Data1.Recordset.FindFirst "Id = " & StrSearch
If Data1.Recordset.NoMatch Then
    MsgBox "·„ Ì „ «ÌÃ«œ «·”›—… «·„ÿ·Ê»… Ì„ﬂ‰ﬂ «ÌÃ«œ «·”›—… »«·«‰ ﬁ«· ⁄»— «·”Ã·« ", vbOKOnly, " ‰»ÌÂ"
End If
End Sub

