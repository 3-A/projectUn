VERSION 5.00
Begin VB.Form Form3_1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��� ������ �������"
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
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "HotelID"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   360
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   480
      Width           =   9375
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      DataField       =   "HotelName"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   360
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1320
      Width           =   9375
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      DataField       =   "HotelAdrees"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   360
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   2160
      Width           =   9375
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      DataField       =   "HotelMark"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   360
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   3000
      Width           =   9375
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      DataField       =   "NoOfRoom"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   360
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   3840
      Width           =   9375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   495
      Left            =   5348
      TabIndex        =   5
      Top             =   7440
      Width           =   1215
   End
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "��� ������:"
      Height          =   375
      Index           =   0
      Left            =   10080
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "��� ������:"
      Height          =   375
      Index           =   1
      Left            =   10080
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "����� ������:"
      Height          =   375
      Index           =   2
      Left            =   10080
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "���� ������:"
      Height          =   375
      Index           =   3
      Left            =   10080
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "��� �����:"
      Height          =   375
      Index           =   4
      Left            =   10080
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   3840
      Width           =   1575
   End
End
Attribute VB_Name = "Form3_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Unload Form3_1
End Sub

Private Sub Data1_Reposition()
Data1.Caption = "����� " & CStr(Data1.Recordset.AbsolutePosition + 1) & " �� " & CStr(Data1.Recordset.RecordCount)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form3.Show
End Sub
Private Sub Form_Activate()
Data1.Recordset.MoveLast
Data1.Recordset.MoveFirst
Data1.Recordset.FindFirst "HotelID = " & StrSearch
If Data1.Recordset.NoMatch Then
    MsgBox "�� ��� ����� ������ ������� ����� ����� ������ ��������� ��� �������", vbOKOnly, "�����"
End If
End Sub

