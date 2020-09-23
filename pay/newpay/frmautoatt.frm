VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmautoatt 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1575
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4320
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   4320
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   3480
      Top             =   600
   End
   Begin MSComctlLib.ProgressBar pg 
      Height          =   255
      Left            =   953
      TabIndex        =   0
      Top             =   840
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000010&
      Caption         =   "Genrating Attendence"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000016&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   4335
   End
End
Attribute VB_Name = "frmautoatt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' this module is used to generate automatic payroll

Dim temp As New ADODB.Recordset
Dim rsesi As New ADODB.Recordset

Private Sub Form_Activate()
'cn.Execute("insert into atten
Timer1.Enabled = True
End Sub

Private Sub Form_Load()
OPENdatabase
Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
If temp.State = adStateOpen Then temp.Close
temp.Open "select * from tempatd where month = '" & Format(dt, "dd-mmm-yy") & "'", cn, adOpenKeyset, adLockOptimistic
If temp.EOF = False Then
    YesNo = MsgBox("Attendance of this month already generated if you" & vbCrLf & "regenrate, this will affect your Advance Master" & vbCrLf & "Do you want to re-generate?", vbQuestion + vbYesNo)
    If YesNo = vbNo Then
        Unload Me
        frmatn.Show
    Else
        End If
    
 End If
     cn.Execute ("delete from tempatd")
pg = pg + 10
pg = pg + 10
'mpf = epf
cn.Execute ("insert into tempatd (t_no,day_wk,day_pay,month,basic,pbasic,da,pvda,hra,lic,clbal,plbal,epf,conv,spall) (select t_no," & days & " as day_wk," & days & " as day_pay,'" & Format(dt, "dd-mmm-yy") & "' as month,basic, basic as pbasic," & vda & " as da," & vda & " as pvda,hra,lic,cl as clbal,pl as plbal,((basic+" & vda & ")*(" & epf & " /100)) as epf,convy as conv,spall  from empmast)")
'((basic + da + hra) * (1.75 / 100)) AS ESI

pg = pg + 10
If temp.State = adStateOpen Then temp.Close
temp.Open "select * from tempatd", cn, adOpenStatic, adLockOptimistic
'MsgBox temp(0)
cn.Execute ("UPDATE TEMPATD a SET ESI = (select ((Pbasic + PVda + hra+conv+SPALL) * (esi / 100)) from empmast b where a.t_no = b.t_no)")
cn.Execute ("update tempatd set epf = Round(epf, 0)")
cn.Execute ("update tempatd set esi = 0 where (Pbasic + PVda + hra+conv+SPALL) >= " & esicut & "")
pg = pg + 10
If temp.State = adStateOpen Then temp.Close
temp.Open "select * from advlon", cn, adOpenKeyset, adLockOptimistic
While Not temp.EOF
If temp!adv_amt < temp!per_month Then
cn.Execute ("update advlon set per_month = adv_amt where t_no = " & temp!T_NO & "")
End If
temp.MoveNext
Wend



cn.Execute ("UPDATE TEMPATD a  SET ADVREC = (select per_month from advlon b where b.t_no = a.t_no)")
cn.Execute ("UPDATE TEMPATD a  SET ADVREC = 0 where advrec IS NULL")

pg = pg + 10
Set rsesi = cn.Execute("select t_no,esi from tempatd")
Dim nesi As Double

While Not rsesi.EOF
nesi = ROUNDUP(rsesi!esi)
cn.Execute ("update tempatd set esi = " & nesi & " where t_no = " & rsesi!T_NO & "")
rsesi.MoveNext
nesi = 0
Wend

pg = pg + 30

'/// total update
Dim totearn As Double
Dim totdedn As Double
If temp.State = adStateOpen Then temp.Close
temp.Open "select * from tempatd", cn, adOpenStatic, adLockOptimistic
While temp.EOF = False
txtpbasic = temp!pbasic
txtpvda = temp!pvda
txtotamt = IIf(IsNull(temp!otamt), 0, temp!otamt)
txthra = IIf(IsNull(temp!HRA), 0, temp!HRA)
txtconv = IIf(IsNull(temp!conv), 0, temp!conv)
txtarrpay = IIf(IsNull(temp!arrpay), 0, temp!arrpay)
txtspall = IIf(IsNull(temp!SPALL), 0, temp!SPALL)
txtEPF = IIf(IsNull(temp!epf), 0, temp!epf)
txtlic = IIf(IsNull(temp!LIC), 0, temp!LIC)
txtadvrec = IIf(IsNull(temp!advrec), 0, temp!advrec)
txtloan = IIf(IsNull(temp!loan), 0, temp!loan)
txtotded = IIf(IsNull(temp!otded), 0, temp!otded)
'// new entry one line
txtesic = IIf(IsNull(temp!esi), 0, temp!esi)
totearn = 0
totdedn = 0
totearn = txtpbasic + txtpvda + txtotamt + txthra + txtconv + txtarrpay + txtspall
totdedn = txtEPF + txtlic + txtadvrec + txtloan + txtotded + txtesic
'txttotal.Text = totearn - totdedn
cn.Execute ("update tempatd set total = (" & totearn & " - " & totdedn & " ) where t_no = " & temp!T_NO & " ")
temp.MoveNext
Wend
pg = pg + 20

Unload Me
'MsgBox "ATTENDENCE GENERETED", vbInformation
frmatn.Show
'frmaten.Show
End Sub

