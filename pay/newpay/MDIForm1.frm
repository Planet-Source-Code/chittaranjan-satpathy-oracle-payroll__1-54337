VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Payroll Management System V1.0"
   ClientHeight    =   4515
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6615
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnumaster 
      Caption         =   "Master"
      Begin VB.Menu mnuempmast 
         Caption         =   "Employee Master"
      End
      Begin VB.Menu mnuadvmast 
         Caption         =   "Advance Master"
      End
      Begin VB.Menu mnuset 
         Caption         =   "Setting"
      End
   End
   Begin VB.Menu mnuatten 
      Caption         =   "Attendence"
      Begin VB.Menu mnuatn 
         Caption         =   "Attendence Entry"
      End
      Begin VB.Menu mnuautogen 
         Caption         =   "Auto Generete"
      End
      Begin VB.Menu line3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuupdate 
         Caption         =   "Update Attendance"
      End
   End
   Begin VB.Menu mnurpt 
      Caption         =   "Report"
      Begin VB.Menu mnuatdchk 
         Caption         =   "Attendence Checkprint"
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu mnusum 
         Caption         =   "Summary"
      End
      Begin VB.Menu mnupaureg 
         Caption         =   "Payment Register"
      End
      Begin VB.Menu mnupayslip 
         Caption         =   "Payment Slip"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuepfdedn 
         Caption         =   "EPF Deduction"
      End
      Begin VB.Menu mnulicdedn 
         Caption         =   "LIC Deduction"
      End
      Begin VB.Menu mnupfm 
         Caption         =   "PFM Report"
      End
      Begin VB.Menu mnuesic 
         Caption         =   "ESIC Deduction"
      End
      Begin VB.Menu mnuadvdedn 
         Caption         =   "Advance Deduction"
      End
      Begin VB.Menu mnudp 
         Caption         =   "Payment Slip Report"
      End
   End
   Begin VB.Menu mnuutil 
      Caption         =   "Utility"
      Begin VB.Menu mnulogin 
         Caption         =   "Login"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuABOUT 
         Caption         =   "About !"
      End
      Begin VB.Menu MNUTEST 
         Caption         =   "TEST"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'     PAYROLL SYSTEM  - by CHITTARANJAN SATPATHY
' THIS IS A COMPLETE PAYROLL PROGRAM IN INDIAN STYLE TO
' CALCULATE PAYROLL FOR INDUSTRIAL PURPOSE.
'
' FIRST YOU NEED TO CREATE ORACLE DATABASE
' TO RUN THIS PROGRAM A SEPARATE README.TXT IS PROVIDED
' FOR CREATING USER AND TABLES IN ORACLE PLEASE FOLLOW
' THOSE INSTRUCTIONS.
' IF YOU LIKE THIS PROGRAM PLEASE DON'T FORGET
' VOTE FOR ME
' mail me for ur sugession at satpathy_cor@yahoo.co.in
'******************************************************

Dim temp As New ADODB.Recordset

Private Sub MDIForm_Load()
' connect to database using ADO

OPENdatabase
End Sub

Private Sub mnuABOUT_Click()
frmabout.Show vbModal
End Sub

Private Sub mnuadvdedn_Click()
frmpayslip.Caption = "Advance Report"
frmpayslip.Show

End Sub

Private Sub mnuadvmast_Click()
frmadvlon.Show
End Sub

Private Sub mnuatdchk_Click()
frmpayslip.Caption = "Attendance Checkprint Report"
frmpayslip.Show
End Sub

Private Sub mnuatn_Click()
'frmaten.Show
If temp.State = adStateOpen Then temp.Close
temp.Open "select * from tempatd", cn, adOpenKeyset, adLockOptimistic
If temp.EOF = True Then
MsgBox "You Need to Generate Attendance", vbInformation
Exit Sub
End If
frmatn.Show
End Sub

Private Sub mnuautogen_Click()
frmmonthyr.Show
End Sub

Private Sub mnudp_Click()
frmpayslip.Show
frmpayslip.Caption = "Payment Slip"

End Sub

Private Sub mnuempmast_Click()
frmempmast.Show
End Sub

Private Sub mnuepfdedn_Click()
frmpayslip.Caption = "EPF Report"
frmpayslip.Show
End Sub

Private Sub mnuesic_Click()
frmpayslip.Caption = "ESI Deduction Report"
frmpayslip.Show
'ChngPrinterOrientationPortrait Me
'p = InputBox("Enter Month Date", "Esi Deduction")
'p = Format(p, "DD-MMM-YYYY")
'
'If temp.State = adStateOpen Then temp.Close
'temp.Open "SELECT EMPMAST.NAME, ATTD.T_NO, ATTD.Month ," _
'    & "(ATTD.PBASIC+ATTD.PVDA) as bval , ATTD.ESI, ATTD.ESIC " _
'    & " From EMPMAST, ATTD where attd.t_no = empmast.t_no" _
'    & " AND MONTH = '" & p & "'", cn, adOpenKeyset, adLockOptimistic
'Set dresi.DataSource = temp
'dresi.Show

End Sub

Private Sub mnuexit_Click()
Unload Me
End Sub

Private Sub mnulicdedn_Click()
frmpayslip.Caption = "LIC Deduction Report"
frmpayslip.Show
End Sub

Private Sub mnulogin_Click()
frmLogin.Show
End Sub

Private Sub mnupaureg_Click()
frmpayslip.Show

frmpayslip.Caption = "Payment Register"

End Sub

Private Sub mnupayslip_Click()
ChngPrinterOrientationPortrait Me
p = InputBox("ENTER DATE OF PAYROLL", "PAYSLIP")
p = Format(p, "DD-MMM-YYYY")

If temp.State = adStateOpen Then temp.Close
temp.Open "SELECT EMPMAST.NAME AS EXPR1, ATTD.T_NO, ATTD.MONTH," _
    & "ATTD.DAY_WK, ATTD.DAY_PAY, ATTD.OTHRS, ATTD.SPALL," _
    & "ATTD.CL, ATTD.PL, ATTD.LWP, ATTD.ARRPAY, ATTD.OTDED," _
    & "ATTD.ORDED, ATTD.ABDAY, ATTD.ADVREC, ATTD.LOAN," _
    & " ATTD.BASIC, ATTD.PBASIC, ATTD.PVDA, ATTD.OTAMT," _
    & "ATTD.HRA, ATTD.CONV, ATTD.EPF, ATTD.LIC, ATTD.CLBAL," _
    & "ATTD.PLBAL, ATTD.TOTAL, ATTD.ESI, ATTD.DA, ATTD.ESIC," _
    & " (ATTD.PBASIC+ATTD.PVDA+ATTD.SPALL+ATTD.ARRPAY+ATTD.OTAMT+ATTD.HRA+ATTD.CONV) AS TOTER," _
    & " (ATTD.EPF+ATTD.ADVREC+ATTD.LOAN+ATTD.LIC+ATTD.ESI+ATTD.OTDED) AS TOTDD," _
    & "ATTD.ROWID From EMPMAST, ATTD WHERE EMPMAST.T_NO = ATTD.T_NO AND MONTH = '" & p & "'", cn, adOpenKeyset, adLockOptimistic
Set DataReport1.DataSource = temp
DataReport1.Show




End Sub

Private Sub mnupfm_Click()
ChngPrinterOrientationLandscape Me
p = InputBox("ENTER DATE OF PAYROLL", "PAYSLIP")
p = Format(p, "DD-MMM-YYYY")
If temp.State = adStateOpen Then temp.Close
temp.Open "SELECT SUM(PBASIC + pvda) AS bval, SUM(epf) AS pepf," _
     & "SUM(epf - ((pbasic + pvda) * .0833)) AS eshare, " _
     & " SUM(pbasic + pvda) * (1.1 / 100) AS adm, SUM(pbasic + pvda)" _
     & "* (.01 / 100) AS adm1, SUM(PBASIC + pvda)*.0833 AS eps From ATTD where MONTH = '" & p & "'", cn, adOpenKeyset, adLockOptimistic
Dim TEMP1 As New ADODB.Recordset
If TEMP1.State = adStateOpen Then TEMP1.Close
TEMP1.Open "SELECT COUNT(*) From ATTD where MONTH = '" & p & "'", cn, adOpenKeyset, adLockOptimistic


Set PFMREPO.DataSource = temp
PFMREPO.Sections("SECTION1").Controls("VAR1").Caption = TEMP1(0)
PFMREPO.Show
End Sub

Private Sub mnuset_Click()
frmsetting.Show
End Sub

Private Sub mnusum_Click()
frmpayslip.Caption = "Summary Report"
frmpayslip.Show
End Sub

Private Sub mnuupdate_Click()
frmupdate.Show
End Sub
