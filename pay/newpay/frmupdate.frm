VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmupdate 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1335
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3810
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   3810
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   2760
      Top             =   240
   End
   Begin MSComctlLib.ProgressBar pg 
      Height          =   255
      Left            =   638
      TabIndex        =   0
      Top             =   720
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Updating"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1358
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmupdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim temp As New ADODB.Recordset

Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Timer1_Timer()
If temp.State = adStateOpen Then temp.Close
temp.Open "Select * from tempatd", cn, adOpenKeyset, adLockOptimistic

If temp.EOF = True Then
MsgBox "No Attendance found to update, or records already updated ", vbInformation
GoTo msg
End If

If MsgBox("Once Records updated you cannot change attendance. Are you Sure?", vbQuestion + vbYesNo) = vbYes Then

cn.Execute ("insert into attd select * from tempatd")
pg = pg + 30
cn.Execute ("update empmast a set cl = cl - (select cl from tempatd b where a.t_no = b.t_no) where t_no in (select t_no from tempatd)")
pg = pg + 35
cn.Execute ("update empmast a set pl = pl - (select pl from tempatd b where a.t_no = b.t_no) where t_no in (select t_no from tempatd)")
pg = pg + 30
cn.Execute ("update advlon a set adv_amt = adv_amt - (select advrec from tempatd b where a.t_no = b.t_no) where t_no in (select t_no from tempatd)")
cn.Execute ("delete from tempatd")
pg = pg + 5
End If

msg:
Unload Me
End Sub
