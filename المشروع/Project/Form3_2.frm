VERSION 5.00
Begin VB.Form Form3_2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Õ–› »Ì«‰«  «·›‰«œﬁ"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11910
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   3  'Windows Default
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
      RecordSource    =   "Hotel"
      RightToLeft     =   -1  'True
      Top             =   8175
      Width           =   11910
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Œ—ÊÃ"
      Height          =   495
      Left            =   5475
      TabIndex        =   5
      Top             =   7320
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      DataField       =   "NoOfRoom"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   480
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   3720
      Width           =   9375
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      DataField       =   "HotelMark"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   480
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   2880
      Width           =   9375
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      DataField       =   "HotelAdrees"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   480
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   2040
      Width           =   9375
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      DataField       =   "HotelName"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   480
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1200
      Width           =   9375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "HotelID"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   480
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   9375
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "⁄œœ «·€—›:"
      Height          =   375
      Index           =   4
      Left            =   10200
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "œ—Ã… «·›‰œﬁ:"
      Height          =   375
      Index           =   3
      Left            =   10200
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "⁄‰Ê«‰ «·›‰œﬁ:"
      Height          =   375
      Index           =   2
      Left            =   10200
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "«”„ «·›‰œﬁ:"
      Height          =   375
      Index           =   1
      Left            =   10200
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "—„“ «·›‰œﬁ:"
      Height          =   375
      Index           =   0
      Left            =   10200
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "Form3_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Form3_2
End Sub

Private Sub Data1_Reposition()
Data1.Caption = "«·”Ã· " & CStr(Data1.Recordset.AbsolutePosition + 1) & " „‰ " & CStr(Data1.Recordset.RecordCount)
End Sub

Private Sub Form_Activate()
Data1.Recordset.MoveLast
Data1.Recordset.MoveFirst
Data1.Recordset.FindFirst "HotelId = " & StrSearch
If Data1.Recordset.NoMatch Then
    MsgBox "·„ Ì „ «ÌÃ«œ «·›‰œﬁ «·„ÿ·Ê» Ì„ﬂ‰ﬂ «ÌÃ«œ «·›‰œﬁ »«·«‰ ﬁ«· ⁄»— «·”Ã·« ", vbOKOnly, " ‰»ÌÂ"
Else
     If MsgBox("Â· «‰  „ «ﬂœ „‰ Õ–› «·”Ã·", vbYesNo + vbQuestion, "ÿ·» «· «ﬂÌœ") = vbNo Then Exit Sub
     Data1.Recordset.Delete
     Data1.Recordset.MoveFirst
End If

End Sub

Private Sub Form_Load()
'Form2_0.Show 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form3.Show
End Sub


