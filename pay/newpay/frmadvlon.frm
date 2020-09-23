VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmadvlon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Advance Master"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   6735
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3015
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   3135
      Begin MSMask.MaskEdBox txtadv_date 
         Height          =   375
         Left            =   1560
         TabIndex        =   11
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd-mmm-yy"
         Mask            =   "##-##-####"
         PromptChar      =   "_"
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&CLose"
         Height          =   375
         Left            =   1560
         TabIndex        =   10
         Top             =   2280
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Save"
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox txtPER_MONTH 
         DataField       =   "PER_MONTH"
         DataMember      =   "Command1"
         DataSource      =   " "
         Height          =   405
         Left            =   1440
         TabIndex        =   7
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtADV_AMT 
         DataField       =   "ADV_AMT"
         DataMember      =   "Command1"
         DataSource      =   " "
         Height          =   405
         Left            =   360
         TabIndex        =   5
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtT_NO 
         DataField       =   "T_NO"
         DataMember      =   "Command1"
         DataSource      =   " "
         Height          =   405
         Left            =   360
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   1920
         TabIndex        =   8
         Top             =   360
         Width           =   435
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Deduction"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   1800
         TabIndex        =   6
         Top             =   1080
         Width           =   915
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Advance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   600
         TabIndex        =   4
         Top             =   1080
         Width           =   810
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Ticket No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   480
         TabIndex        =   2
         Top             =   360
         Width           =   915
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   2910
      Left            =   3360
      Style           =   1  'Simple Combo
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   360
      Width           =   3135
   End
End
Attribute VB_Name = "frmadvlon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim temp As New ADODB.Recordset
Dim tmp() As String

Private Sub Combo1_Click()
tmp = Split(Combo1.Text, ":")
txtT_NO = tmp(0)
If temp.State = adStateOpen Then temp.Close
temp.Open "select * from advlon where T_no = " & txtT_NO & "", cn, adOpenKeyset, adLockOptimistic
If Not temp.EOF Then
Command1.Caption = "&Save"
txtadv_date = Format(temp!adv_date, "dd-mm-yyyy")
txtADV_AMT = temp!adv_amt
txtPER_MONTH.Text = temp!per_month
Else
Command1.Caption = "&Add"
txtadv_date = Format(Date, "dd-mm-yyyy")
txtADV_AMT = ""
txtPER_MONTH.Text = ""
End If
End Sub

Private Sub Command1_Click()
If temp.State = adStateOpen Then temp.Close
temp.Open "select * from advlon where t_no = " & txtT_NO & "", cn, adOpenKeyset, adLockOptimistic
If temp.EOF = True Then
cn.Execute "insert into advlon values(" & txtT_NO & "," & txtADV_AMT & ",0," & txtPER_MONTH & ",0,'" & Format(txtadv_date, "dd-mmm-yyyy") & "')"
Command1.Caption = "&Save"
Else
cn.Execute "update advlon set ADV_AMT = " & txtADV_AMT & ",per_month = " & txtPER_MONTH & ", adv_date = '" & Format(txtadv_date, "dd-mmm-yyyy") & "' where t_no = " & txtT_NO & ""

End If

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
Combo1.Text = ""
If temp.State = adStateOpen Then temp.Close
temp.Open "select T_no,name from empmast", cn, adOpenKeyset, adLockOptimistic
While Not temp.EOF
Combo1.AddItem temp(0) & ":" & temp(1)
temp.MoveNext
Wend
End Sub
