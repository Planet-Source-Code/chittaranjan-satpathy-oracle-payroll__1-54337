VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmaten 
   Caption         =   "Attendence Entry"
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   9315
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   4575
      Left            =   2880
      ScaleHeight     =   4515
      ScaleWidth      =   3435
      TabIndex        =   32
      Top             =   960
      Visible         =   0   'False
      Width           =   3495
      Begin VB.CommandButton Command3 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   1800
         TabIndex        =   35
         Top             =   3960
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Ok"
         Height          =   375
         Left            =   480
         TabIndex        =   34
         Top             =   3960
         Width           =   855
      End
      Begin VB.ListBox List1 
         Height          =   3765
         Left            =   120
         TabIndex        =   33
         Top             =   120
         Width           =   3135
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Search"
      Height          =   375
      Left            =   5280
      TabIndex        =   31
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   3975
      Left            =   360
      TabIndex        =   15
      Top             =   720
      Width           =   4575
      Begin VB.TextBox txtT_NO 
         DataField       =   "T_NO"
         DataMember      =   " "
         DataSource      =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtDAY_WK 
         Alignment       =   1  'Right Justify
         DataField       =   "DAY_WK"
         DataMember      =   " "
         DataSource      =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   240
         TabIndex        =   1
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtDAY_PAY 
         Alignment       =   1  'Right Justify
         DataField       =   "DAY_PAY"
         DataMember      =   " "
         DataSource      =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1620
         TabIndex        =   2
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtOTHRS 
         Alignment       =   1  'Right Justify
         DataField       =   "OTHRS"
         DataMember      =   " "
         DataSource      =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   240
         TabIndex        =   5
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox txtSPALL 
         Alignment       =   1  'Right Justify
         DataField       =   "SPALL"
         DataMember      =   " "
         DataSource      =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1620
         TabIndex        =   6
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox txtCL 
         Alignment       =   1  'Right Justify
         DataField       =   "CL"
         DataMember      =   " "
         DataSource      =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3000
         TabIndex        =   3
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtPL 
         Alignment       =   1  'Right Justify
         DataField       =   "PL"
         DataMember      =   " "
         DataSource      =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3720
         TabIndex        =   4
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtARRPAY 
         Alignment       =   1  'Right Justify
         DataField       =   "ARRPAY"
         DataMember      =   " "
         DataSource      =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1620
         TabIndex        =   9
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox txtOTDED 
         Alignment       =   1  'Right Justify
         DataField       =   "OTDED"
         DataMember      =   " "
         DataSource      =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3000
         TabIndex        =   13
         Top             =   3420
         Width           =   1335
      End
      Begin VB.TextBox txtADVREC 
         Alignment       =   1  'Right Justify
         DataField       =   "ADVREC"
         DataMember      =   " "
         DataSource      =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   240
         TabIndex        =   8
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox txtLOAN 
         Alignment       =   1  'Right Justify
         DataField       =   "LOAN"
         DataMember      =   " "
         DataSource      =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3000
         TabIndex        =   7
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox txtHRA 
         Alignment       =   1  'Right Justify
         DataField       =   "HRA"
         DataMember      =   " "
         DataSource      =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   240
         TabIndex        =   11
         Top             =   3420
         Width           =   1335
      End
      Begin VB.TextBox txtCONV 
         Alignment       =   1  'Right Justify
         DataField       =   "CONV"
         DataMember      =   " "
         DataSource      =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1620
         TabIndex        =   12
         Top             =   3420
         Width           =   1335
      End
      Begin VB.TextBox txtLIC 
         Alignment       =   1  'Right Justify
         DataField       =   "LIC"
         DataMember      =   " "
         DataSource      =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3000
         TabIndex        =   10
         Top             =   2640
         Width           =   1335
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   495
         Left            =   960
         TabIndex        =   36
         Top             =   480
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   873
         _Version        =   327681
         Enabled         =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "Employee Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1440
         TabIndex        =   30
         Top             =   480
         Width           =   2655
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
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   915
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Days Worked"
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
         Index           =   2
         Left            =   240
         TabIndex        =   28
         Top             =   975
         Width           =   1245
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Days Payable"
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
         Left            =   1560
         TabIndex        =   27
         Top             =   960
         Width           =   1290
      End
      Begin VB.Label d 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "OT Hrs"
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
         Left            =   240
         TabIndex        =   26
         Top             =   1680
         Width           =   645
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "SPL Allowance"
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
         Index           =   5
         Left            =   1680
         TabIndex        =   25
         Top             =   1680
         Width           =   1350
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "CL"
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
         Index           =   6
         Left            =   3240
         TabIndex        =   24
         Top             =   960
         Width           =   240
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "PL"
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
         Index           =   7
         Left            =   3960
         TabIndex        =   23
         Top             =   960
         Width           =   240
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Arear Pay"
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
         Index           =   9
         Left            =   1680
         TabIndex        =   22
         Top             =   2400
         Width           =   900
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Other Deduction"
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
         Index           =   10
         Left            =   3000
         TabIndex        =   21
         Top             =   3120
         Width           =   1440
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
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
         Index           =   12
         Left            =   240
         TabIndex        =   20
         Top             =   2400
         Width           =   810
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Loan"
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
         Index           =   13
         Left            =   3240
         TabIndex        =   19
         Top             =   1680
         Width           =   450
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "House Rent"
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
         Index           =   18
         Left            =   240
         TabIndex        =   18
         Top             =   3120
         Width           =   1065
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Convaynce"
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
         Index           =   19
         Left            =   1680
         TabIndex        =   17
         Top             =   3120
         Width           =   1020
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "LIC"
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
         Index           =   21
         Left            =   3480
         TabIndex        =   16
         Top             =   2400
         Width           =   285
      End
   End
   Begin MSFlexGridLib.MSFlexGrid fg1 
      Height          =   3855
      Left            =   5040
      TabIndex        =   14
      Top             =   840
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   6800
      _Version        =   393216
      Rows            =   12
      FixedCols       =   0
      FormatString    =   "PARTICULARS               |     AMOUNT           "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmaten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim temp As New ADODB.Recordset
Dim emp() As String
Dim rsemp As New ADODB.Recordset
Dim rsempdt As New ADODB.Recordset
Private Sub Command1_Click()

Set rsemp = cn.Execute("select t_no,name from empmast")
List1.Clear
While rsemp.EOF = False
List1.AddItem rsemp(0) & ":" & rsemp(1)
rsemp.MoveNext
Wend
Picture1.Visible = True
End Sub

Private Sub Command2_Click()
Picture1.Visible = False
emp = Split(List1.Text, ":", 2)
Set rsempdt = cn.Execute("select * from tempatd where t_no = " & emp(0) & "")
SETFIELDS

End Sub

Private Sub Command3_Click()
Picture1.Visible = False
End Sub

Private Sub Form_Load()
fg1.TextMatrix(1, 0) = "Month"
fg1.TextMatrix(2, 0) = "Basic Rate"
fg1.TextMatrix(3, 0) = "Paid Basic"
fg1.TextMatrix(4, 0) = "VDA"
fg1.TextMatrix(5, 0) = "Paid VDA"
fg1.TextMatrix(6, 0) = "EPF"
fg1.TextMatrix(7, 0) = "ESI"
fg1.TextMatrix(8, 0) = "OT Amount"
fg1.TextMatrix(9, 0) = "CL Balance"
fg1.TextMatrix(10, 0) = "PL Balance"
fg1.TextMatrix(11, 0) = "Total Amount"
OPENdatabase
If temp.State = adStateOpen Then temp.Close
temp.Open "select max(t_no) from tempatd", cn, adOpenKeyset, adLockOptimistic
UpDown1.Max = temp(0)
UpDown1.Min = 1
txtT_NO.Text = UpDown1.Min
End Sub

Private Sub SETFIELDS()
'Label1.Caption = emp(1)
If temp.State = adStateOpen Then temp.Close
temp.Open "select name from empmast where t_no = " & txtT_NO.Text & "", cn, adOpenKeyset, adLockOptimistic
Label1.Caption = temp(0)
txtT_NO = rsempdt!T_NO
txtDAY_WK = rsempdt!day_wk
txtDAY_PAY = rsempdt!day_pay
txtCL = rsempdt!CL
txtPL = rsempdt!PL
txtOTHRS = rsempdt!othrs
txtADVREC = IIf(IsNull(rsempdt!advrec), 0, rsempdt!advrec)
txtSPALL = IIf(IsNull(rsempdt!SPALL), 0, rsempdt!SPALL)
txtARRPAY = rsempdt!arrpay
txtCONV = rsempdt!conv
txtOTDED = rsempdt!otded
txtLOAN = rsempdt!loan
txtLIC = rsempdt!LIC
txtHRA = rsempdt!HRA
fg1.TextMatrix(1, 1) = rsempdt!Month
fg1.TextMatrix(2, 1) = rsempdt!BASIC
fg1.TextMatrix(3, 1) = rsempdt!pbasic
fg1.TextMatrix(4, 1) = rsempdt!DA
fg1.TextMatrix(5, 1) = rsempdt!pvda
fg1.TextMatrix(6, 1) = rsempdt!epf
fg1.TextMatrix(7, 1) = rsempdt!esi
fg1.TextMatrix(8, 1) = rsempdt!otamt
fg1.TextMatrix(9, 1) = rsempdt!clbal
fg1.TextMatrix(10, 1) = rsempdt!plbal
'fg1.TextMatrix(11, 0) = "Total Amount"
End Sub

Private Sub txtT_NO_Change()
Set rsempdt = Nothing
Set rsempdt = cn.Execute("select * from tempatd where t_no = " & txtT_NO.Text & " ")
If Not rsempdt.EOF Then
SETFIELDS
Else
MsgBox "Employee not found", vbInformation
End If
End Sub

Private Sub UpDown1_Change()
txtT_NO = UpDown1.Value
End Sub
