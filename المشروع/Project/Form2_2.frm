VERSION 5.00
Begin VB.Form Form2_2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«œŒ«· »Ì«‰«  «·„”«›—Ì‰"
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
   Begin VB.CommandButton Command1 
      Caption         =   "Œ—ÊÃ"
      Height          =   495
      Left            =   5400
      TabIndex        =   10
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Data Data1 
      Align           =   2  'Align Bottom
      Caption         =   "«·”Ã·«  "
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
      RecordSource    =   "Traveller"
      RightToLeft     =   -1  'True
      Top             =   8175
      Width           =   11910
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "PersonID"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   120
      Width           =   9375
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      DataField       =   "PersonName"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   960
      Width           =   9375
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      DataField       =   "PersonAddress"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1800
      Width           =   9375
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      DataField       =   "TripId"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   2640
      Width           =   9375
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      DataField       =   "HotelId"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   3480
      Width           =   9375
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "—„“ «·‘Œ’:"
      Height          =   375
      Index           =   0
      Left            =   10080
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "«”„ «·‘Œ’:"
      Height          =   375
      Index           =   1
      Left            =   10080
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "«·⁄‰Ê«‰:"
      Height          =   375
      Index           =   2
      Left            =   10080
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "—„“ «·”›—…:"
      Height          =   375
      Index           =   3
      Left            =   10080
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "—„“ «·›‰œﬁ:"
      Height          =   375
      Index           =   4
      Left            =   10080
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   3480
      Width           =   1575
   End
End
Attribute VB_Name = "Form2_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.Recordset.Update
Unload Form2_2
End Sub

Private Sub Form_Activate()
Data1.Recordset.AddNew

End Sub

Private Sub Form_Unload(Cancel As Integer)
Form2.Show
End Sub

Private Sub Data1_Reposition()
Data1.Caption = "«·”Ã· " & CStr(Data1.Recordset.AbsolutePosition + 1) & " „‰ " & CStr(Data1.Recordset.RecordCount)
End Sub


