VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmatn 
   Caption         =   "Employee Attendence Entry"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7590
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame6 
      Height          =   615
      Left            =   2040
      TabIndex        =   63
      Top             =   5880
      Width           =   4935
      Begin VB.CommandButton Command1 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1210
         TabIndex        =   65
         Top             =   180
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   64
         Top             =   180
         Width           =   1095
      End
   End
   Begin MSFlexGridLib.MSFlexGrid fg1 
      Height          =   2055
      Left            =   533
      TabIndex        =   56
      Top             =   3620
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   3625
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
      BackColor       =   16777215
      ForeColorFixed  =   0
      ForeColorSel    =   16777215
      BackColorBkg    =   -2147483633
      GridColor       =   0
      HighLight       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid fg3 
      Height          =   2415
      Left            =   5453
      TabIndex        =   61
      Top             =   1200
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   4260
      _Version        =   393216
      Rows            =   7
      FixedCols       =   0
      GridColor       =   -2147483626
      FormatString    =   "Detail                | Value                "
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
   Begin MSFlexGridLib.MSFlexGrid fg2 
      Height          =   4455
      Left            =   8340
      TabIndex        =   55
      Top             =   1200
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   7858
      _Version        =   393216
      Rows            =   12
      FixedCols       =   0
      GridColor       =   -2147483626
      ScrollBars      =   0
      FormatString    =   "Head               |     Amount"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame5 
      Caption         =   "Deduction"
      Height          =   1335
      Left            =   9173
      TabIndex        =   50
      Top             =   3368
      Visible         =   0   'False
      Width           =   2055
      Begin VB.TextBox txtepf 
         DataField       =   "epf"
         DataMember      =   "Command1"
         Height          =   285
         Left            =   1080
         TabIndex        =   52
         Top             =   240
         Width           =   840
      End
      Begin VB.TextBox txtesic 
         DataField       =   "esic"
         DataMember      =   "Command1"
         Height          =   285
         Left            =   1080
         TabIndex        =   51
         Top             =   840
         Width           =   840
      End
      Begin VB.Label Label11 
         Caption         =   "EPF"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "ESIC"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   840
         Width           =   615
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Absent days"
      Height          =   1095
      Left            =   9173
      TabIndex        =   45
      Top             =   2288
      Visible         =   0   'False
      Width           =   2055
      Begin VB.TextBox txtabday 
         DataField       =   "abday"
         DataMember      =   "Command1"
         Height          =   285
         Left            =   1080
         TabIndex        =   47
         Top             =   255
         Width           =   840
      End
      Begin VB.TextBox txtlwp 
         DataField       =   "lwp"
         DataMember      =   "Command1"
         Height          =   285
         Left            =   1080
         TabIndex        =   46
         Top             =   600
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Leave Without Pay"
         Height          =   495
         Index           =   7
         Left            =   120
         TabIndex        =   49
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Absent Days"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   270
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Earnings"
      Height          =   2415
      Left            =   6293
      TabIndex        =   32
      Top             =   2288
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox txtpbasic 
         DataField       =   "pbasic"
         DataMember      =   "Command1"
         Height          =   285
         Left            =   1215
         TabIndex        =   38
         Top             =   240
         Width           =   1320
      End
      Begin VB.TextBox txtotamt 
         DataField       =   "otamt"
         DataMember      =   "Command1"
         Height          =   285
         Left            =   1200
         TabIndex        =   37
         Top             =   930
         Width           =   1320
      End
      Begin VB.TextBox txthra 
         DataField       =   "hra"
         DataMember      =   "Command1"
         Height          =   285
         Left            =   1200
         TabIndex        =   36
         Top             =   1275
         Width           =   1320
      End
      Begin VB.TextBox txtconv 
         DataField       =   "conv"
         DataMember      =   "Command1"
         Height          =   285
         Left            =   1200
         TabIndex        =   35
         Top             =   1620
         Width           =   1320
      End
      Begin VB.TextBox txttotal 
         DataField       =   "total"
         DataMember      =   "Command1"
         Height          =   285
         Left            =   1200
         TabIndex        =   34
         Top             =   1965
         Width           =   1320
      End
      Begin VB.TextBox txtpvda 
         Height          =   285
         Left            =   1200
         TabIndex        =   33
         Top             =   585
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Paid Basic"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   255
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Paid VDA"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   564
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "OT Amount"
         Height          =   375
         Left            =   120
         TabIndex        =   42
         Top             =   873
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "HRA"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   1302
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "CONV"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   1611
         Width           =   735
      End
      Begin VB.Label Label16 
         Caption         =   "TOTAL"
         Height          =   375
         Left            =   120
         TabIndex        =   39
         Top             =   1920
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Attendence "
      Height          =   2535
      Left            =   533
      TabIndex        =   13
      Top             =   1100
      Width           =   4935
      Begin VB.TextBox txtlic 
         DataField       =   "lic"
         DataMember      =   "Command1"
         Height          =   340
         Left            =   3480
         TabIndex        =   59
         Top             =   1440
         Width           =   840
      End
      Begin VB.TextBox txtDAY_WK 
         DataField       =   "DAY_WK"
         DataMember      =   "Command1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   57
         Top             =   840
         Width           =   1200
      End
      Begin VB.TextBox txtadvrec 
         DataField       =   "advrec"
         DataMember      =   "Command1"
         Height          =   340
         Left            =   120
         TabIndex        =   27
         Top             =   2040
         Width           =   1080
      End
      Begin VB.TextBox txtloan 
         DataField       =   "loan"
         DataMember      =   "Command1"
         Height          =   340
         Left            =   1320
         TabIndex        =   26
         Top             =   2040
         Width           =   1080
      End
      Begin VB.TextBox txtotded 
         DataField       =   "otded"
         DataMember      =   "Command1"
         Height          =   340
         Left            =   2520
         TabIndex        =   25
         Top             =   2040
         Width           =   1080
      End
      Begin VB.TextBox txtarrpay 
         DataField       =   "arrpay"
         DataMember      =   "Command1"
         Height          =   340
         Left            =   2280
         TabIndex        =   24
         Top             =   1440
         Width           =   1080
      End
      Begin VB.TextBox txtspall 
         DataField       =   "spall"
         DataMember      =   "Command1"
         Height          =   340
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   1080
      End
      Begin VB.TextBox txtothrs 
         DataField       =   "othrs"
         DataMember      =   "Command1"
         Height          =   340
         Left            =   1200
         TabIndex        =   17
         Top             =   1440
         Width           =   1080
      End
      Begin VB.TextBox txtDAY_PAY 
         DataField       =   "DAY_PAY"
         DataMember      =   "Command1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1320
         TabIndex        =   16
         Top             =   840
         Width           =   1200
      End
      Begin VB.TextBox txtcl 
         DataField       =   "cl"
         DataMember      =   "Command1"
         Height          =   345
         Left            =   2520
         TabIndex        =   15
         Top             =   840
         Width           =   840
      End
      Begin VB.TextBox txtpl 
         DataField       =   "pl"
         DataMember      =   "Command1"
         Height          =   345
         Left            =   3480
         TabIndex        =   14
         Top             =   840
         Width           =   840
      End
      Begin VB.Label Label19 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   720
         TabIndex        =   62
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label12 
         Caption         =   "LIC"
         Height          =   255
         Left            =   3720
         TabIndex        =   60
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Arrear Pay"
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
         Index           =   8
         Left            =   2400
         TabIndex        =   31
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Other Dedn."
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
         Index           =   9
         Left            =   2520
         TabIndex        =   30
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Advance Rec"
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
         Left            =   240
         TabIndex        =   29
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label4 
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
         Height          =   255
         Left            =   1680
         TabIndex        =   28
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Days Paid"
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
         Index           =   2
         Left            =   1440
         TabIndex        =   23
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "O.T.Hrs"
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
         Index           =   3
         Left            =   1320
         TabIndex        =   22
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Spcl. Allown."
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
         Index           =   4
         Left            =   120
         TabIndex        =   21
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "CL "
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
         Index           =   5
         Left            =   2760
         TabIndex        =   20
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   6
         Left            =   3720
         TabIndex        =   19
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Days Wrkd"
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
         TabIndex        =   58
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Employee Information"
      Height          =   1095
      Left            =   1013
      TabIndex        =   0
      Top             =   1088
      Visible         =   0   'False
      Width           =   7695
      Begin VB.TextBox txtclbal 
         DataField       =   "clbal"
         DataMember      =   "Command1"
         Enabled         =   0   'False
         Height          =   345
         Left            =   4800
         TabIndex        =   10
         Top             =   240
         Width           =   600
      End
      Begin VB.TextBox txtplbal 
         DataField       =   "plbal"
         DataMember      =   "Command1"
         Enabled         =   0   'False
         Height          =   345
         Left            =   6360
         TabIndex        =   9
         Top             =   240
         Width           =   600
      End
      Begin VB.TextBox txtmonth 
         DataField       =   "month"
         DataMember      =   "Command1"
         Enabled         =   0   'False
         Height          =   345
         Left            =   2880
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtempno 
         DataField       =   "empno"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1080
         TabIndex        =   5
         Top             =   240
         Width           =   720
      End
      Begin VB.TextBox txtvda 
         DataField       =   "vda"
         DataMember      =   "Command1"
         Enabled         =   0   'False
         Height          =   345
         Left            =   2880
         TabIndex        =   2
         Top             =   675
         Width           =   960
      End
      Begin VB.TextBox txtBasic 
         DataField       =   "Basic"
         DataMember      =   "Command1"
         Enabled         =   0   'False
         Height          =   345
         Left            =   1080
         TabIndex        =   1
         Top             =   675
         Width           =   1080
      End
      Begin VB.Label Label14 
         Caption         =   "CL Balance"
         Height          =   255
         Left            =   3960
         TabIndex        =   12
         Top             =   285
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "PL Balance"
         Height          =   255
         Left            =   5400
         TabIndex        =   11
         Top             =   285
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Ticket No."
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   285
         Width           =   1335
      End
      Begin VB.Label Label17 
         Caption         =   "MONTH"
         Height          =   255
         Left            =   2160
         TabIndex        =   7
         Top             =   285
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Basic"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label18 
         Caption         =   "VDA"
         Height          =   255
         Left            =   2280
         TabIndex        =   3
         Top             =   720
         Width           =   375
      End
   End
   Begin VB.Label Label20 
      BackColor       =   &H80000011&
      Caption         =   "Attendance Entry Module"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000016&
      Height          =   615
      Left            =   0
      TabIndex        =   66
      Top             =   0
      Width           =   12000
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   5
      Height          =   6015
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   840
      Width           =   11415
   End
End
Attribute VB_Name = "frmatn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsemp As New ADODB.Recordset
Dim rsatd As New ADODB.Recordset
Dim rsadv As New ADODB.Recordset
Dim totrec As Integer
Dim temp As New ADODB.Recordset
Dim rstemp As New ADODB.Recordset
Dim UPDFLAG As Boolean
'Const epf = 12
'Const esi = 1.75


Private Sub Command1_Click()
If MsgBox("Are you Sure", vbQuestion + vbYesNo) = vbYes Then
Unload Me
End If

End Sub

Private Sub Command2_Click()
'On Error GoTo mess

Dim sql As String
txtothrs.Text = IIf(txtothrs.Text = "", 0, txtothrs.Text)
txtspall.Text = IIf(txtspall.Text = "", 0, txtspall.Text)
txtcl.Text = IIf(txtcl.Text = "", 0, txtcl.Text)
txtpl.Text = IIf(txtpl.Text = "", 0, txtpl.Text)
txtarrpay.Text = IIf(txtarrpay.Text = "", 0, txtarrpay.Text)
txtabday.Text = IIf(txtabday.Text = "", 0, txtabday.Text)
txtadvrec.Text = IIf(txtadvrec.Text = "", 0, txtadvrec.Text)
txtloan.Text = IIf(txtloan.Text = "", 0, txtloan.Text)
txtotamt.Text = IIf(txtotamt.Text = "", 0, txtotamt.Text)
txtlic.Text = IIf(txtlic.Text = "", 0, txtlic.Text)
txtclbal.Text = IIf(txtclbal.Text = "", 0, txtclbal.Text)
txtplbal.Text = IIf(txtplbal.Text = "", 0, txtplbal.Text)
findtotal

sql = "INSERT INTO tempatd ( empno, DAY_WK, DAY_PAY, othrs, spall, cl, pl, lwp, arrpay, otded, abday, advrec, loan, Basic, pbasic, pvda, otamt, hra, conv, epf, lic, esic, clbal, plbal,month ) values (" & txtempno.Text & "," & txtDAY_WK.Text & "," & txtDAY_PAY.Text & "," & txtothrs.Text & "," & txtspall.Text & "," & txtcl.Text & "," & txtpl.Text & "," & txtlwp.Text & "," & txtarrpay.Text & "," & txtabday.Text & "," & txtadvrec.Text & "," & txtloan.Text & "," & txtBasic.Text & "," & txtpbasic.Text & "," & txtpvda.Text & "," & txtotamt.Text & "," & txthra.Text & "," & txtconv.Text & "," & txtepf.Text & "," & txtlic.Text & "," & txtesic.Text & "," & txtclbal.Text & "," & txtplbal.Text & ",'" & txtmonth.Text & "'"

'////////// hjkhkhkl
cn.Execute ("UPDATE tempatd SET  total = " & txttotal.Text & ",DAY_WK=" & txtDAY_WK.Text & ", DAY_PAY=" & txtDAY_PAY.Text & ", othrs=" & txtothrs.Text & ", spall=" & txtspall.Text & ", cl=" & txtcl.Text & ", pl=" & txtpl.Text & ", lwp=" & txtlwp.Text & ", arrpay=" & txtarrpay.Text & ", abday=" & txtabday.Text & ", advrec=" & txtadvrec.Text & ", loan=" & txtloan.Text & ", Basic=" & txtBasic.Text & ", pbasic=" & txtpbasic.Text & ", pvda=" & txtpvda.Text & ", otamt=" & txtotamt.Text & ", hra=" & txthra.Text & ", conv=" & txtconv.Text & ", epf=" & txtepf.Text & ", lic=" & txtlic.Text & ", esi=" & txtesic.Text & ", clbal=" & txtclbal.Text & ", plbal=" & txtplbal.Text & "  where t_no=" & txtempno.Text & " ")
MsgBox "Record Updated", vbInformation, "Information"
fg1.CellForeColor = vbRed
Exit Sub

mess:
MsgBox "Unable to process your request", vbCritical

End Sub




Private Sub Command3_Click()




End Sub

Private Sub fg1_Click()
fg1.Col = 1
SETFIELDS
End Sub

Private Sub Form_Activate()
MDIForm1.mnuatten.Enabled = False
End Sub

Private Sub Form_Load()
'Dim dt As Integer

If days = 0 Then
If temp.State = adStateOpen Then temp.Close
temp.Open "select * from tempatd", cn, adOpenKeyset, adLockOptimistic
mnth = temp!Month
dt = temp!Month
Label20.Caption = "Attendence Entry for the Month of " & MonthName(Month(dt)) & "," & Year(dt)
Select Case Month(temp!Month)
Case 1, 3, 5, 7, 8, 10, 12
days = 31
Case 4, 6, 9, 11
days = 30
Case 2
If Int(Year(Date) / 4) * 4 = Year(Date) Then
days = 29
Else
days = 28
End If
End Select
End If


'frmatn.Top = 500
'frmatn.Left = 500
With fg2
.Row = 1: .Col = 0: .Text = "Paid Basic"
.Row = 2: .Col = 0: .Text = "Vda"
.Row = 3: .Col = 0: .Text = "OT Amount"
.Row = 4: .Col = 0: .Text = "HRA"
.Row = 5: .Col = 0: .Text = "Conv"
.Row = 6: .Col = 0: .Text = "EPF"
.Row = 7: .Col = 0: .Text = "LIC"
.Row = 8: .Col = 0: .Text = "ESIC"
.Row = 9: .Col = 0: .Text = "Absent Days"
.Row = 10: .Col = 0: .Text = "Leave W/P"
.Row = 11: .Col = 0: .Text = "TOTAL"
End With
With fg3
.Row = 1: .Col = 0: .Text = "Month"
.Row = 2: .Col = 0: .Text = "Ticket No."
.Row = 3: .Col = 0: .Text = "Basic"
.Row = 4: .Col = 0: .Text = "VDA"
.Row = 5: .Col = 0: .Text = "PL Bal."
.Row = 6: .Col = 0: .Text = "CL Bal"
'.Row = 7: .Col = 0: .Text = "LIC"
'.Row = 8: .Col = 0: .Text = "ESIC"
'.Row = 9: .Col = 0: .Text = "Absent Days"
'.Row = 10: .Col = 0: .Text = "Leave W/P"
'.Row = 11: .Col = 0: .Text = "TOTAL"
End With
fg1.Cols = 10

With fg1
.Row = 0: .Col = 0: .Text = "Serial": .ColWidth(0) = 800
.Row = 0: .Col = 1: .Text = "Ticket No.": .ColWidth(1) = 800
.Row = 0: .Col = 2: .Text = "Employee Name": .ColWidth(2) = 3000
.Row = 0: .Col = 3: .Text = "Basic": .ColWidth(3) = 800
.Row = 0: .Col = 4: .Text = "CL BAL": .ColWidth(4) = 800
.Row = 0: .Col = 5: .Text = "PL BAL": .ColWidth(5) = 800
.Row = 0: .Col = 6: .Text = "Grade": .ColWidth(6) = 800
.Row = 0: .Col = 7: .Text = "HRA": .ColWidth(7) = 800
.Row = 0: .Col = 8: .Text = "Conv": .ColWidth(8) = 800
.Row = 0: .Col = 9: .Text = "LIC"

End With
Dim sql As String
sql = "select count(*) from empmast"
If temp.State = adStateOpen Then temp.Close
temp.Open sql, cn, adOpenKeyset, adLockOptimistic
If temp.EOF = False Then
fg1.Rows = temp(0) + 5
fg1.Cols = 10
End If
Set rsemp = cn.Execute("select * from empmast order by T_no")
fg1.Row = 0
fg1.Col = 0
SL = 0
While rsemp.EOF = False
SL = SL + 1
fg1.Row = fg1.Row + 1
fg1.Col = 0
fg1.Text = SL

fg1.Col = 1
fg1.Text = rsemp("t_no")
fg1.Col = 2
fg1.Text = rsemp("name")
fg1.Col = 3
fg1.Text = rsemp("BASIC")
fg1.Col = 4
fg1.Text = rsemp("CL")
fg1.Col = 5
fg1.Text = rsemp("PL")
fg1.Col = 6
'fg1.Text = IIf(rsemp("grade") = "A", "Staff", IIf(rsemp("grade") = "B", "Security", IIf(rsemp("grade") = "C", "Labour", "Unknown")))
fg1.Col = 7
fg1.Text = rsemp!HRA
fg1.Col = 8
fg1.Text = rsemp!convy
fg1.Col = 9
fg1.Text = IIf(IsNull(rsemp("lic")), 0, rsemp("lic"))
rsemp.MoveNext
Wend
fg1.Row = 1
fg1_Click
End Sub

Private Sub SETFIELDS()
On Error GoTo msg

If rstemp.State = adStateOpen Then
rstemp.Close
End If

Set rstemp = cn.Execute("select * from TEMPATD where T_no = " & fg1.TextMatrix(fg1.Row, 1) & "")
txtempno.Text = fg1.TextMatrix(fg1.Row, 1)

txtDAY_WK.Text = rstemp("day_wk")
txtDAY_PAY.Text = rstemp("day_pay")
txtBasic.Text = rstemp("BASIC")
txtpbasic.Text = rstemp("PBASIC")
txtpvda.Text = rstemp("PVDA")
txtmonth.Text = dt
txtvda.Text = vda
txtcl.Text = rstemp("CL")
txtpl.Text = rstemp("PL")
txtplbal.Text = rstemp("PLBAL")
txtclbal.Text = rstemp("CLBAL")
txtarrpay.Text = rstemp("ARRPAY")
txtspall.Text = IIf(IsNull(rstemp("spall")), 0, rstemp!SPALL)
txtadvrec.Text = IIf(IsNull(rstemp("ADVREC")), 0, rstemp!advrec)
txtothrs.Text = rstemp("OTHRS")
txtloan.Text = rstemp("LOAN")
txtlic.Text = rstemp!LIC

'End If
With fg3
.TextMatrix(1, 1) = txtmonth
.TextMatrix(2, 1) = txtempno
.TextMatrix(3, 1) = txtBasic
.TextMatrix(4, 1) = txtvda
.TextMatrix(5, 1) = txtplbal
.TextMatrix(6, 1) = txtclbal


End With

txtDAY_WK_Change
Exit Sub
msg:

End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIForm1.mnuatten.Enabled = True
End Sub

Private Sub txtADVREC_Change()
daychange
End Sub

Private Sub txtARRPAY_Change()
daychange
End Sub

Private Sub txtday_pay_Change()
On Error GoTo msg
txtDAY_WK_Change
msg:
End Sub

Private Sub txtDAY_WK_Change()
daychange
End Sub

Private Sub txtLIC_Change()
daychange

End Sub

Private Sub txtothrs_Change()
On Error GoTo msg
txtDAY_WK_Change
msg:
End Sub

Sub daychange()
On Error GoTo msg
txtpbasic.Text = Round(txtBasic.Text / days * txtDAY_PAY.Text, 2)
fg2.TextMatrix(1, 1) = Format(txtpbasic.Text, "######.#") 'basic rate
txtabday.Text = days - IIf(IsEmpty(txtDAY_WK.Text), 0, txtDAY_WK.Text)
txtpvda.Text = Round(vda / days * IIf(IsEmpty(txtDAY_PAY.Text), 0, txtDAY_PAY.Text), 2)
fg2.TextMatrix(2, 1) = Format(txtpvda.Text, "######.#") 'pvda
txtlwp.Text = days - txtDAY_PAY.Text
txtconv.Text = Round((conv / days) * txtDAY_WK.Text, 2)
'fg1.TextMatrix(fg1.Row, 8)

If (Val(txtpbasic.Text) + Val(txtpvda.Text) + Val(txtotamt.Text) + Val(txthra.Text) + Val(txtconv.Text)) >= esicut Then
txtesic.Text = 0
Else
txtesic.Text = ROUNDUP((Val(txtpbasic.Text) + Val(txtpvda.Text) + Val(txtotamt.Text) + Val(txthra.Text) + Val(txtconv.Text)) * (esi / 100))
End If

If IsEmpty(txtothrs.Text) Or txtothrs.Text = "" Then
txtotamt.Text = 0
fg2.TextMatrix(3, 1) = 0
Else
txtotamt.Text = Round(((txtBasic.Text + vda) / 240) * txtothrs.Text, 2) 'overtime hours
fg2.TextMatrix(3, 1) = Format(txtotamt.Text, "######.00")
End If
'txtSPALL.Text = IIf(fg1.TextMatrix(fg1.Row, 6) = "Labour", 50, 0)
txtepf.Text = Round((Val(txtpbasic.Text) + Val(txtpvda.Text)) * (epf / 100), 0)
fg2.TextMatrix(6, 1) = Format(txtepf.Text, "######.00") 'epf amount
txthra.Text = fg1.TextMatrix(fg1.Row, 7)
                '= "Yes", Round(txtBasic.Text * (hra / 100), 0), 0)
fg2.TextMatrix(7, 1) = Format(txtlic.Text, "######.00") 'lic amount

fg2.TextMatrix(4, 1) = Format(txthra.Text, "######.00") 'hra amount
txtconv.Text = FormatNumber((fg1.TextMatrix(fg1.Row, 8) / days) * txtDAY_WK, 2)
fg2.TextMatrix(5, 1) = Format(txtconv.Text, "######.00") 'convaynce
If (Val(txtpbasic.Text) + Val(txtpvda.Text) + Val(txtotamt.Text) + Val(txthra.Text) + Val(txtconv.Text)) >= esicut Then
txtesic.Text = 0
Else
txtesic.Text = ROUNDUP((Val(txtpbasic.Text) + Val(txtpvda.Text) + Val(txtotamt.Text) + Val(txthra.Text) + Val(txtconv.Text)) * (esi / 100))
End If

fg2.TextMatrix(8, 1) = Format(txtesic.Text, "######.00") 'esic amount
'txtLIC.Text = fg1.TextMatrix(fg1.Row, 9)
fg2.TextMatrix(9, 1) = Format(txtabday.Text, "######.00") 'absent days
fg2.TextMatrix(10, 1) = Format(txtlwp.Text, "######.00") 'leave without payment
fg1.Col = 2
Label19.Caption = fg1.Text
findtotal

msg:

End Sub

Sub findtotal()
Dim totearn As Double
Dim totdedn As Double
totearn = Val(txtpbasic.Text) + Val(txtpvda) + Val(txtotamt) + Val(txthra) + Val(txtconv) + Val(txtarrpay) + Val(txtspall)
totdedn = Val(txtepf) + Val(txtlic) + Val(txtadvrec) + Val(txtloan) + Val(txtotded) + Val(txtesic)
txttotal.Text = totearn - totdedn
fg2.TextMatrix(11, 1) = txttotal
End Sub

Private Sub txtSPALL_Change()
daychange
End Sub
