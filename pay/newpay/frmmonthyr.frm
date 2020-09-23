VERSION 5.00
Begin VB.Form frmmonthyr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Month and Year"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3675
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   3675
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtyr 
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Top             =   600
      Width           =   855
   End
   Begin VB.ComboBox cmbmonth 
      Height          =   315
      Left            =   1320
      TabIndex        =   3
      Top             =   240
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      Height          =   135
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   3735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Month Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmmonthyr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim temp As New ADODB.Recordset
Private Sub Command1_Click()
Dim mnstr As Integer
dt = CDate(cmbmonth & "-" & txtyr)
Select Case cmbmonth.ListIndex + 1
Case 1, 3, 5, 7, 8, 10, 12
days = 31
Case 4, 6, 9, 11
days = 30
Case 2
If Int(txtyr.Text / 4) * 4 = txtyr.Text Then
days = 29
Else
days = 28
End If
End Select
If temp.State = adStateOpen Then temp.Close
temp.Open "select * from attd where month = '" & Format(dt, "dd-mmm-yy") & "'", cn, adOpenKeyset, adLockOptimistic

If temp.EOF = False Then
MsgBox "Attendance May Already Updated!", vbCritical
Exit Sub
End If
Unload Me
frmautoatt.Show

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
cmbmonth.Clear
cmbmonth.AddItem "January"
cmbmonth.AddItem "February"
cmbmonth.AddItem "March"
cmbmonth.AddItem "April"
cmbmonth.AddItem "May"
cmbmonth.AddItem "June"
cmbmonth.AddItem "July"
cmbmonth.AddItem "August"
cmbmonth.AddItem "September"
cmbmonth.AddItem "October"
cmbmonth.AddItem "November"
cmbmonth.AddItem "December"
cmbmonth.ListIndex = Month(Date) - 1
txtyr.Text = Year(Date)
End Sub

