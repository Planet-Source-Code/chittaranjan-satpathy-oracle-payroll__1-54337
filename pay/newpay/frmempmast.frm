VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmempmast 
   Caption         =   "Employee Master"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7590
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture3 
      Height          =   4935
      Left            =   2513
      Picture         =   "frmempmast.frx":0000
      ScaleHeight     =   4875
      ScaleWidth      =   3555
      TabIndex        =   59
      Top             =   1028
      Width           =   3615
      Begin VB.CommandButton Command7 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   1920
         TabIndex        =   62
         Top             =   4440
         Width           =   735
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&Ok"
         Height          =   375
         Left            =   600
         TabIndex        =   61
         Top             =   4440
         Width           =   855
      End
      Begin VB.ListBox List1 
         Height          =   3570
         Left            =   360
         TabIndex        =   60
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Employee Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   840
         TabIndex        =   63
         Top             =   240
         Width           =   1935
      End
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   1313
      Top             =   3188
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=MSDAORA.1;Password=challan;User ID=challan;Data Source=oracor;Persist Security Info=True"
      OLEDBString     =   "Provider=MSDAORA.1;Password=challan;User ID=challan;Data Source=oracor;Persist Security Info=True"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Height          =   1935
      Left            =   8513
      TabIndex        =   54
      Top             =   4268
      Width           =   1815
      Begin VB.TextBox txtreason 
         Height          =   735
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   24
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtleftdate 
         Height          =   285
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Reason"
         Height          =   195
         Index           =   4
         Left            =   480
         TabIndex        =   56
         Top             =   720
         Width           =   555
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Left Service Date"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   55
         Top             =   120
         Width           =   1245
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   8513
      ScaleHeight     =   2745
      ScaleWidth      =   1425
      TabIndex        =   49
      Top             =   1148
      Width           =   1455
      Begin VB.CommandButton Command5 
         Caption         =   "&Search"
         Height          =   375
         Left            =   120
         TabIndex        =   58
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Close"
         Height          =   375
         Left            =   120
         TabIndex        =   57
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Add"
         Height          =   375
         Left            =   120
         TabIndex        =   50
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   8513
      ScaleHeight     =   2745
      ScaleWidth      =   1425
      TabIndex        =   51
      Top             =   1148
      Width           =   1455
      Begin VB.CommandButton Command3 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   120
         TabIndex        =   53
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Save"
         Height          =   375
         Left            =   120
         TabIndex        =   52
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2415
      Left            =   2678
      TabIndex        =   42
      Top             =   1028
      Width           =   5655
      Begin VB.TextBox txtNAME 
         DataField       =   "NAME"
         DataMember      =   " "
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   480
         Width           =   3615
      End
      Begin VB.TextBox txtT_NO 
         DataField       =   "T_NO"
         DataMember      =   " "
         Enabled         =   0   'False
         Height          =   285
         Left            =   360
         TabIndex        =   0
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtFNAME 
         DataField       =   "FNAME"
         DataMember      =   " "
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   840
         Width           =   3600
      End
      Begin VB.TextBox txtADD1 
         DataField       =   "ADD1"
         DataMember      =   " "
         Height          =   285
         Left            =   1560
         TabIndex        =   43
         Top             =   1200
         Width           =   3600
      End
      Begin VB.TextBox txtADD2 
         DataField       =   "ADD2"
         DataMember      =   " "
         Height          =   285
         Left            =   1560
         TabIndex        =   3
         Top             =   1560
         Width           =   3600
      End
      Begin VB.TextBox txtADD3 
         DataField       =   "ADD3"
         DataMember      =   " "
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Top             =   1920
         Width           =   3600
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "T/No."
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   47
         Top             =   240
         Width           =   435
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Name of the Employee"
         Height          =   195
         Index           =   14
         Left            =   2640
         TabIndex        =   46
         Top             =   240
         Width           =   1605
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fathers Name"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   45
         Top             =   840
         Width           =   990
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Address"
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   44
         Top             =   1200
         Width           =   570
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   2678
      TabIndex        =   25
      Top             =   3428
      Width           =   5655
      Begin VB.TextBox txtspall 
         DataField       =   "LIC"
         DataMember      =   " "
         DataSource      =   " "
         Height          =   285
         Left            =   1440
         TabIndex        =   22
         Top             =   2775
         Width           =   1335
      End
      Begin VB.TextBox txtconvy 
         DataField       =   "DOJ"
         DataMember      =   " "
         DataSource      =   " "
         Height          =   285
         Left            =   120
         TabIndex        =   21
         Top             =   2775
         Width           =   1155
      End
      Begin VB.TextBox txtDESIGNATION 
         DataField       =   "DESIGNATION"
         DataMember      =   " "
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1155
      End
      Begin VB.TextBox txtBOB 
         DataField       =   "BOB"
         DataMember      =   " "
         DataSource      =   " "
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   1560
         Width           =   1155
      End
      Begin VB.TextBox txtDOJ 
         DataField       =   "DOJ"
         DataMember      =   " "
         DataSource      =   " "
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   2160
         Width           =   1155
      End
      Begin VB.TextBox txtSEX 
         DataField       =   "SEX"
         DataMember      =   " "
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   375
         Width           =   1155
      End
      Begin VB.TextBox txtCATEGORY 
         DataField       =   "CATEGORY"
         DataMember      =   " "
         DataSource      =   " "
         Height          =   285
         Left            =   4200
         TabIndex        =   20
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox txtPFNO 
         DataField       =   "PFNO"
         DataMember      =   " "
         Height          =   285
         Left            =   2760
         TabIndex        =   11
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtESINO 
         DataField       =   "ESINO"
         DataMember      =   " "
         DataSource      =   " "
         Height          =   285
         Left            =   2760
         TabIndex        =   15
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox txtLICNO 
         DataField       =   "LICNO"
         DataMember      =   " "
         DataSource      =   " "
         Height          =   285
         Left            =   2760
         TabIndex        =   19
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox txtPL 
         DataField       =   "PL"
         DataMember      =   " "
         Height          =   285
         Left            =   4200
         TabIndex        =   12
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtCL 
         DataField       =   "CL"
         DataMember      =   " "
         DataSource      =   " "
         Height          =   285
         Left            =   4200
         TabIndex        =   16
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox txtBASIC 
         DataField       =   "BASIC"
         DataMember      =   " "
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   372
         Width           =   1335
      End
      Begin VB.TextBox txtDA 
         DataField       =   "DA"
         DataMember      =   " "
         Height          =   285
         Left            =   2760
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtHRA 
         DataField       =   "HRA"
         DataMember      =   " "
         Height          =   285
         Left            =   4230
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtPF 
         DataField       =   "PF"
         DataMember      =   " "
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Top             =   966
         Width           =   1335
      End
      Begin VB.TextBox txtESI 
         DataField       =   "ESI"
         DataMember      =   " "
         DataSource      =   " "
         Height          =   285
         Left            =   1440
         TabIndex        =   14
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txtLIC 
         DataField       =   "LIC"
         DataMember      =   " "
         DataSource      =   " "
         Height          =   285
         Left            =   1440
         TabIndex        =   18
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Special Allowance"
         Height          =   195
         Index           =   23
         Left            =   1560
         TabIndex        =   68
         Top             =   2520
         Width           =   1305
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Convyance"
         Height          =   195
         Index           =   22
         Left            =   315
         TabIndex        =   67
         Top             =   2520
         Width           =   810
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Designation"
         Height          =   195
         Index           =   5
         Left            =   270
         TabIndex        =   41
         Top             =   720
         Width           =   840
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Date of Birth"
         Height          =   195
         Index           =   6
         Left            =   255
         TabIndex        =   40
         Top             =   1305
         Width           =   885
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Date of Join"
         Height          =   195
         Index           =   7
         Left            =   270
         TabIndex        =   39
         Top             =   1905
         Width           =   855
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         Height          =   195
         Index           =   21
         Left            =   555
         TabIndex        =   38
         Top             =   120
         Width           =   270
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Category"
         Height          =   195
         Index           =   8
         Left            =   4680
         TabIndex        =   37
         Top             =   1920
         Width           =   630
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "PF No."
         Height          =   195
         Index           =   9
         Left            =   3120
         TabIndex        =   36
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ESI No."
         Height          =   195
         Index           =   10
         Left            =   3240
         TabIndex        =   35
         Top             =   1320
         Width           =   555
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "LIC No."
         Height          =   195
         Index           =   11
         Left            =   3240
         TabIndex        =   34
         Top             =   1920
         Width           =   540
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "PL Balance"
         Height          =   195
         Index           =   12
         Left            =   4560
         TabIndex        =   33
         Top             =   720
         Width           =   825
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "CL Balance"
         Height          =   195
         Index           =   13
         Left            =   4560
         TabIndex        =   32
         Top             =   1320
         Width           =   825
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Basic Rate"
         Height          =   195
         Index           =   15
         Left            =   1680
         TabIndex        =   31
         Top             =   120
         Width           =   780
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "D.A."
         Height          =   195
         Index           =   16
         Left            =   3240
         TabIndex        =   30
         Top             =   120
         Width           =   315
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "HRA"
         Height          =   195
         Index           =   17
         Left            =   4800
         TabIndex        =   29
         Top             =   120
         Width           =   345
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "PF Dedn. Rate"
         Height          =   195
         Index           =   18
         Left            =   1560
         TabIndex        =   28
         Top             =   720
         Width           =   1065
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ESI Dedn Rate"
         Height          =   195
         Index           =   19
         Left            =   1560
         TabIndex        =   27
         Top             =   1305
         Width           =   1080
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "LIC Deduction"
         Height          =   195
         Index           =   20
         Left            =   1560
         TabIndex        =   26
         Top             =   1905
         Width           =   1020
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   8513
      ScaleHeight     =   2745
      ScaleWidth      =   1425
      TabIndex        =   64
      Top             =   1148
      Width           =   1455
      Begin VB.CommandButton Command10 
         Caption         =   "&Update"
         Height          =   375
         Left            =   120
         TabIndex        =   66
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command9 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   120
         TabIndex        =   65
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   5
      Height          =   5775
      Left            =   2033
      Shape           =   4  'Rounded Rectangle
      Top             =   908
      Width           =   8535
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000010&
      Caption         =   "     Employee Master Module"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000016&
      Height          =   615
      Left            =   0
      TabIndex        =   48
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "frmempmast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rscust As New ADODB.Recordset
Dim temp As New ADODB.Recordset
Dim rsemp As New ADODB.Recordset


Private Sub Command1_Click()
Picture1.Visible = False
Picture2.Visible = True
If temp.State = adStateOpen Then temp.Close
temp.Open "select max(t_no) from empmast", cn, adOpenKeyset, adLockOptimistic
If Not temp.EOF Then
txtT_NO.Text = IIf(IsNull(temp(0)), 0, temp(0)) + 1
Else
txtT_NO.Text = 1
End If

End Sub

Private Sub Command10_Click()
Picture4.Visible = False
Picture1.Visible = True
cn.Execute ("update empmast set name='" & txtNAME & "', fname='" & txtFNAME & "', add1='" & txtADD1 & "', add2='" & txtADD2 & "', add3='" & txtADD3 & "', designation='" & txtDESIGNATION & "', doj='" & Format(txtDOJ, "dd-mmm-yyyy") & "', bob='" & Format(txtBOB, "dd-mmm-yyyy") & "', category='" & txtCATEGORY & "', pfno='" & txtPFNO & "', esino='" & txtESINO & "', licno='" & txtLICNO & "', pl=" & txtpl & ", cl=" & txtcl & ", basic= " & txtBasic & ", da=" & txtDA & ", hra=" & txthra & ", leftdate='" & txtleftdate & "', reason='" & txtreason & "', pf=" & txtPF & ", esi=" & txtESI & ", lic=" & txtlic & ", sex='" & txtSEX & "' where T_no = " & txtT_NO & "")
End Sub

Private Sub Command2_Click()
'On Error GoTo mistake


Picture2.Visible = False
Picture1.Visible = True
'////// insert value
'cn.Execute ("insert into empmast (NAME,FNAME,ADD1 ,ADD2 ,ADD3 ,DESIGNATION ,DOJ ,BOB ,CATEGORY ,PFNO ,ESINO,LICNO,PL ,CL ,T_NO,Basic,DA ,HRA,leftdate ,reason ,PF ,ESI,LIC,SEX,convy ,SPALL) values('" & txtNAME & "', '" & txtFNAME & "', '" & txtADD1 & "', '" & txtADD2 & "', '" & txtADD3 & "', '" & txtDESIGNATION & "', '" & txtDOJ & "', '" & txtBOB & "', '" & txtCATEGORY & "', '" & txtPFNO & "', '" & txtESINO & "', '" & txtLICNO & "', " & txtPL & ", " & txtCL & ", " & txtT_NO & ", " & txtBasic & ", " & txtDA & ", " & txtHRA & ", '" & txtleftdate & "', '" & txtreason & "', " & txtPF & ", " & txtESI & ", " & txtLIC & ", '" & txtSEX & "'," & txtconvy & "," & txtSPALL & ")")
cn.Execute ("insert into empmast (NAME,FNAME,ADD1 ,ADD2 ,ADD3 ,DESIGNATION ,DOJ ,BOB ,CATEGORY ,PFNO ,ESINO,LICNO,PL ,CL ,T_NO,Basic,DA ,HRA,leftdate ,reason ,PF ,ESI,LIC,SEX,convy ,SPALL) values('" & txtNAME & "','" & txtFNAME & "','" & txtADD1 & "','" & txtADD2 & "','" & txtADD3 & "','" & txtDESIGNATION & "','" & Format(txtDOJ, "dd-mmm-yy") & "','" & txtBOB & "','" & txtCATEGORY & "','" & txtPFNO & "','" & txtESINO & "','" & txtLICNO & "'," & txtpl & "," & txtcl & "," & txtT_NO & "," & txtBasic & "," & txtDA & "," & txthra & ",'" & txtleftdate & "','" & txtreason & "'," & txtPF & "," & txtESI & "," & txtlic & ",'" & txtSEX & "'," & txtconvy & "," & txtspall & ")")
Exit Sub
mistake:
MsgBox "Error No.:" & Err.number & vbCrLf & "Unable to save the desire", vbCritical + vbOKOnly
End Sub

Private Sub Command3_Click()
Picture2.Visible = False
Picture1.Visible = True
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
Picture3.Visible = True
List1.Clear
Set rscust = Nothing
Set rscust = cn.Execute("select name from empmast")
Do While Not rscust.EOF
List1.AddItem rscust(0)
rscust.MoveNext
Loop
End Sub

Private Sub Command6_Click()
Set rsempt = Nothing
Set rsemp = cn.Execute("select * from empmast where name = '" & List1.Text & "'")
SETFIELDS
Picture3.Visible = False
Picture1.Visible = False
Picture2.Visible = False
Picture4.Visible = True
End Sub

Private Sub Command7_Click()
Picture3.Visible = flase
End Sub

Private Sub Command9_Click()
Picture4.Visible = False
Picture1.Visible = True

End Sub

Private Sub Form_Load()
'OPENdatabase
Picture3.Visible = False
End Sub
Private Sub SETFIELDS()
  txtNAME.Text = rsemp!Name
 txtFNAME.Text = rsemp!FNAME
 txtADD1.Text = rsemp!ADD1
 txtADD2.Text = rsemp!ADD2
 txtADD3.Text = rsemp!ADD3
 txtDESIGNATION.Text = rsemp!DESIGNATION
 txtDOJ.Text = rsemp!DOJ
 txtBOB.Text = rsemp!BOB
 txtCATEGORY.Text = rsemp!CATEGORY
 txtPFNO.Text = rsemp!PFNO
 txtESINO.Text = rsemp!ESINO
 
 txtLICNO.Text = IIf(IsNull(rsemp!LICNO), "", rsemp!LICNO)
 txtpl.Text = rsemp!PL
 txtcl.Text = rsemp!CL
 txtT_NO.Text = rsemp!T_NO
 txtBasic.Text = rsemp!BASIC
 txtDA.Text = rsemp!DA
 txthra.Text = rsemp!HRA
 txtleftdate.Text = IIf(IsNull(rsemp!LEFTDATE), "", rsemp!LEFTDATE)
 txtreason.Text = IIf(IsNull(rsemp!REASON), "", rsemp!REASON)
 txtPF.Text = rsemp!PF
 txtESI.Text = rsemp!esi
 txtlic = rsemp!LIC
 txtSEX = rsemp!SEX
 txtconvy = rsemp!convy
 txtspall = IIf(IsNull(rsemp!SPALL), 0, rsemp!SPALL)
 
End Sub

Private Sub txtNAME_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
