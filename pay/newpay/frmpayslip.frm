VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmpayslip 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payslip"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7620
   ScaleWidth      =   11910
   WindowState     =   2  'Maximized
   Begin VB.ComboBox txtdate 
      Height          =   315
      Left            =   1080
      TabIndex        =   7
      Top             =   360
      Width           =   1455
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   6255
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   11033
      _Version        =   393217
      BackColor       =   16777215
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"frmpayslip.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      Height          =   375
      Left            =   6840
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   375
      Left            =   6840
      TabIndex        =   3
      Top             =   600
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   735
      Left            =   4320
      TabIndex        =   0
      Top             =   120
      Width           =   2055
      Begin VB.OptionButton Option2 
         Caption         =   "File"
         Height          =   255
         Left            =   1080
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Printer"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Label Label1 
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
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   420
      Width           =   735
   End
End
Attribute VB_Name = "frmpayslip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim temp As New ADODB.Recordset
Dim p As String

Private Sub Combo1_Change()

End Sub

Private Sub Command1_Click()

If txtdate = "" Then
MsgBox "Date Should Not be Blank", vbCritical
Exit Sub
End If

p = Format(txtdate, "DD-MMM-YYYY")

If temp.State = adStateOpen Then temp.Close
temp.Open "SELECT EMPMAST.NAME AS EXPR1,empmast.designation as exprd, ATTD.T_NO, ATTD.MONTH," _
    & "ATTD.DAY_WK, ATTD.DAY_PAY, ATTD.OTHRS, ATTD.SPALL," _
    & "ATTD.CL, ATTD.PL, ATTD.LWP, ATTD.ARRPAY, ATTD.OTDED," _
    & "ATTD.ORDED, ATTD.ABDAY, ATTD.ADVREC, ATTD.LOAN," _
    & " ATTD.BASIC, ATTD.PBASIC, ATTD.PVDA, ATTD.OTAMT," _
    & "ATTD.HRA, ATTD.CONV, ATTD.EPF, ATTD.LIC, ATTD.CLBAL," _
    & "ATTD.PLBAL, ATTD.TOTAL, ATTD.ESI, ATTD.DA, ATTD.ESIC," _
    & " (ATTD.PBASIC+ATTD.PVDA+ATTD.SPALL+ATTD.ARRPAY+ATTD.OTAMT+ATTD.HRA+ATTD.CONV) AS TOTER," _
    & " (ATTD.EPF+ATTD.ADVREC+ATTD.LOAN+ATTD.LIC+ATTD.ESI+ATTD.OTDED) AS TOTDD," _
    & "ATTD.ROWID From EMPMAST, ATTD WHERE EMPMAST.T_NO = ATTD.T_NO AND MONTH = '" & p & "'", cn, adOpenKeyset, adLockOptimistic

'///////// Printing Module Starts from Here /////////
Select Case frmpayslip.Caption
Case "Payment Slip"
PRINTSLIP
Case "Payment Register"
printregister
Case "EPF Report"
printpfrepo
Case "ESI Deduction Report"
printesirepo
Case "Advance Report"
PRINTADVANCE
Case "LIC Deduction Report"
printlic
Case "Summary Report"
printsummary
Case "Attendance Checkprint Report"
printcheck
End Select
If Option1.Value = True Then
Shell ("command.com /c type c:\ranjan.txt >prn "), vbHide
End If
End Sub

Sub PRINTSLIP()
Open "C:\RANJAN.TXT" For Output As #1
countrec = 0
HEAD = "PAYSLIP FOR THE MONTH OF " & MonthName(Month(temp!Month)) & " " & Year(temp!Month)
Print #1, Chr(27) & Chr(67) & Chr(0) & Chr(12)
While Not temp.EOF
Print #1,
Print #1, Space((78 - Len("SUMIT ENTERPRISES.")) / 4) & Chr(27) & "W1" & "SUMIT ENTERPRISES." & Chr(27) & "W0"

Print #1, Space((78 - Len(HEAD)) / 2) & UCase(HEAD)

Print #1, String(78, "-")
Print #1, Space(5) & "Employee Name " & Space(14) & "EMPLOYEE NO " & Space(2) & "DAYS WORKED" & Space(2) & "DAYS PAID"
Print #1,
Print #1, Space(5) & temp!expr1 & Space(30 - Len(temp!expr1)) & Str(temp!T_NO) & Space(11) & Str(temp!day_wk) & Space(10) & Str(temp!day_pay)
Print #1, Space(5) & "--------ACTUAL---------------------------LEAVE TAKEN-------------"
'Print #1, String(78, "-")
Print #1, Space(5) & "BASIC RATE " & Space(5) & "DA" & Space(10) & "OVERTIME" & Space(11) & "CL      EL  "
Print #1, Space(5) & Str(Format(temp!BASIC, "####0.00")) & Space(7) & Str(Format(temp!DA, "#####0.00")) & Space(11) & Str(Format(temp!othrs, "####0.00")) & Space(13) & FormatNumber(temp!CL) & Space(5) & FormatNumber(temp!PL)
Print #1, String(78, "-")
Print #1, Space(5) & String(9, "-") & "EARNINGS" & String(9, "-") & Space(10) & String(9, "-") & "DEDUCTIONS" & String(9, "-")
Print #1, Space(5) & "BASIC          " & Space(10 - Len(FormatNumber(temp!pbasic))) & FormatNumber(temp!pbasic) & Space(12) & "EPF     " & Space(18 - Len(FormatNumber(temp!epf))) & FormatNumber(temp!epf)
Print #1, Space(5) & "VDA            " & Space(10 - Len(FormatNumber(temp!pvda))) & FormatNumber(temp!pvda) & Space(12) & "ADVANCE " & Space(18 - Len(FormatNumber(temp!advrec))) & FormatNumber(temp!advrec)
Print #1, Space(5) & "SPECIAL ALLOUN." & Space(10 - Len(FormatNumber(temp!SPALL))) & FormatNumber(temp!SPALL) & Space(12) & "LOAN    " & Space(18 - Len(FormatNumber(temp!loan))) & FormatNumber(temp!loan)
Print #1, Space(5) & "ARREARS        " & Space(10 - Len(FormatNumber(temp!arrpay))) & FormatNumber(temp!arrpay) & Space(12) & "LIC     " & Space(18 - Len(FormatNumber(temp!LIC))) & FormatNumber(temp!LIC)
Print #1, Space(5) & "OVERIMTE       " & Space(10 - Len(FormatNumber(temp!otamt))) & FormatNumber(temp!otamt) & Space(12) & "ESIC    " & Space(18 - Len(FormatNumber(temp!esi))) & FormatNumber(temp!esi)
Print #1, Space(5) & "H.R.A          " & Space(10 - Len(FormatNumber(temp!HRA))) & FormatNumber(temp!HRA) & Space(12) & "OTHERS  " & Space(18 - Len(FormatNumber(temp!otded))) & FormatNumber(temp!otded)
Print #1, Space(5) & "CONVEYANCE     " & Space(10 - Len(FormatNumber(temp!conv))) & FormatNumber(temp!conv)
Print #1, String(78, "-")
Print #1, Space(5) & "TOTAL          " & Space(10 - Len(FormatNumber(temp!toter))) & FormatNumber(temp!toter) & Space(12) & "Total   " & Space(18 - Len(FormatNumber(temp!TOTDD))) & FormatNumber(temp!TOTDD)
Print #1, String(78, "-")
Print #1, Space(5) & "NET PAY        " & Space(10 - Len(FormatNumber(temp!toter - temp!TOTDD))) & FormatNumber(temp!toter - temp!TOTDD) & Space(5) & "Leave Balance CL:" & temp!clbal & Space(5) & "PL:" & temp!PLbal
Print #1, String(78, "-")
Print #1,
Print #1,
Print #1,
Print #1,
Print #1,

countrec = countrec + 1
If countrec >= 2 Then
countrec = 0
Print #1, Chr(12)
End If

temp.MoveNext
Wend
Close #1
RichTextBox1.FileName = "c:\ranjan.txt"

'Unload Me

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Form_Load()
If temp.State = adStateOpen Then temp.Close
temp.Open "select month from attd group by month", cn, adOpenKeyset, adLockOptimistic
While Not temp.EOF
txtdate.AddItem temp(0)
temp.MoveNext
Wend

'txtdate = "01-01-2004"
End Sub

Sub printregister()
Dim nBASIC As Double
Dim nVDA, nOTAMT, nHRA, nCONV, nSPALL As Double
Dim nARRPAY, nGROSS, nEPF, nesi, nLOAN As Double
Dim nADV_REC, nLIC, nTOTAL, nNET As Double

Open "C:\RANJAN.TXT" For Output As #1
countrec = 0
SL = 1
HEAD = "PAYMENT REGISTER FOR THE MONTH OF " & MonthName(Month(temp!Month)) & " " & Year(temp!Month)
Print #1, Chr(27) & Chr(67) & Chr(0) & Chr(12)
Print #1,
Print #1, Space((84 - Len("SUMIT ENTERPRISES.")) / 2) & Chr(27) & "W1" & "SUMIT ENTERPRISES." & Chr(27) & "W0"

Print #1, Space((84 - Len(HEAD))) & UCase(HEAD) & Chr(15)
'Print #1, String(78, "-")
Print #1, "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
Print #1, "SL    EMPLOYEE NAME             BASIC                 ATTENDANCE                                P A Y M E N T S                                               D E D U C T I O N S"
Print #1, "NO    EMPLOYEE NO.                         -------------------------------    ------------------------------------------------------   ----------------------------------------------------------------------"
Print #1, "      DESIGNATION                           PAID    WKD    CL    EL    W/P     BASIC      VDA       OT      HRA     CONV.   SP.ALLW.   ARRS.     GROSS    EPF     ESI      LOAN    ADVANCE    LIC       TOTAL    NET PAY"
Print #1, "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
While Not temp.EOF
Print #1, SL & Space(5 - Len(SL)) & Left(temp!expr1, 24) & Space(25 - Len(Left(temp!expr1, 24))) & Space(8 - Len(FormatNumber(temp!BASIC))) & FormatNumber(temp!BASIC) & Space(10 - Len(FormatNumber(temp!day_wk))) & FormatNumber(temp!day_wk) _
& Space(7 - Len(FormatNumber(temp!day_pay))) & FormatNumber(temp!day_pay) & Space(6 - Len(FormatNumber(temp!CL))) & FormatNumber(temp!CL) & Space(6 - Len(FormatNumber(temp!PL))) & FormatNumber(temp!PL) & Space(6 - Len(FormatNumber(temp!lwp))) & FormatNumber(temp!lwp) _
& Space(11 - Len(FormatNumber(temp!pbasic))) & FormatNumber(temp!pbasic) & Space(10 - Len(FormatNumber(temp!pvda))) & FormatNumber(temp!pvda) & Space(9 - Len(FormatNumber(temp!otamt))) & FormatNumber(temp!otamt) & Space(8 - Len(FormatNumber(temp!HRA))) & FormatNumber(temp!HRA) _
& Space(10 - Len(FormatNumber(temp!conv))) & FormatNumber(temp!conv) & Space(10 - Len(FormatNumber(temp!SPALL))) & FormatNumber(temp!SPALL) & Space(8 - Len(FormatNumber(temp!arrpay))) & FormatNumber(temp!arrpay) & Space(12 - Len(FormatNumber(temp!toter))) & FormatNumber(temp!toter) _
& Space(8 - Len(FormatNumber(temp!epf))) & FormatNumber(temp!epf) & Space(8 - Len(FormatNumber(temp!esi))) & FormatNumber(temp!esi) & Space(9 - Len(FormatNumber(temp!loan))) & FormatNumber(temp!loan) & Space(11 - Len(FormatNumber(temp!advrec))) & FormatNumber(temp!advrec) _
& Space(8 - Len(FormatNumber(temp!LIC))) & FormatNumber(temp!LIC) & Space(11 - Len(FormatNumber(temp!TOTDD))) & FormatNumber(temp!TOTDD) _
& Space(11 - Len(FormatNumber(temp!toter - temp!TOTDD))) & FormatNumber(temp!toter - temp!TOTDD)

Print #1, Space(11) & temp!T_NO
Print #1, Space(5) & temp!exprd
Print #1, "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"

'Print #1, 30 - Len(Left(temp!expr1, 24))

'If countrec >= 2 Then
'countrec = 0
'Print #1, Chr(12)
'End If
nBASIC = nBASIC + temp!pbasic
nVDA = nVDA + temp!pvda
nOTAMT = nOTAMT + temp!otamt
nHRA = nHRA + temp!HRA
nCONV = nCONV + temp!conv
nSPALL = nSPALL + temp!SPALL
nARRPAY = nARRPAY + temp!arrpay
nGROSS = nGROSS + temp!toter
nEPF = nEPF + temp!epf
nesi = nesi + temp!esi
nLOAN = nLOAN + temp!loan
nADV_REC = nADV_REC + temp!advrec
nLIC = nLIC + temp!LIC
nTOTAL = nTOTAL + temp!TOTDD
nNET = nNET + (nGROSS - nTOTAL)
SL = SL + 1

countrec = countrec + 1

temp.MoveNext

Wend
Print #1, Space(72) & Space(12 - Len(FormatNumber(nBASIC))) & FormatNumber(nBASIC) & Space(10 - Len(FormatNumber(nVDA))) & FormatNumber(nVDA) _
& Space(9 - Len(FormatNumber(nOTAMT))) & FormatNumber(nOTAMT) & Space(8 - Len(FormatNumber(nHRA))) & FormatNumber(nHRA) _
& Space(10 - Len(FormatNumber(nCONV))) & FormatNumber(nCONV) & Space(10 - Len(FormatNumber(nSPALL))) & FormatNumber(nSPALL) _
& Space(8 - Len(FormatNumber(nARRPAY))) & FormatNumber(nARRPAY) & Space(12 - Len(FormatNumber(nGROSS))) & FormatNumber(nGROSS) _
& Space(8 - Len(FormatNumber(nEPF))) & FormatNumber(nEPF) & Space(8 - Len(FormatNumber(nesi))) & FormatNumber(nesi) _
& Space(9 - Len(FormatNumber(nLOAN))) & FormatNumber(nLOAN) & Space(11 - Len(FormatNumber(nADV_REC))) & FormatNumber(nADV_REC) _
& Space(8 - Len(FormatNumber(nLIC))) & FormatNumber(nLIC) & Space(11 - Len(FormatNumber(nTOTAL))) & FormatNumber(nTOTAL) _
& Space(11 - Len(FormatNumber(nNET))) & FormatNumber(nNET)
Print #1, "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"

Close #1


RichTextBox1.FileName = "c:\ranjan.txt"

End Sub

Sub printpfrepo()
Dim bval, ps, sBVAL, sPS, sEPF, sFPF As Double
Open "C:\RANJAN.TXT" For Output As #1
countrec = 0
HEAD = "STATEMENT OF PF DEDUCTION FOR THE MONTH OF " & MonthName(Month(temp!Month)) & " " & Year(temp!Month)
Print #1, Chr(27) & Chr(67) & Chr(0) & Chr(12)
Print #1,
Print #1, Space((84 - Len("SUMIT ENTERPRISES.")) / 4) & Chr(27) & "W1" & "SUMIT ENTERPRISES." & Chr(27) & "W0"
Print #1, Space((84 - Len(HEAD)) / 2) & UCase(HEAD)
Print #1,
Print #1, "-------------------------------------------------------------------------------------"
Print #1, "SL. EMP.   EMPLOYEE                   BASIC         P.F.         PENSION        EPF  "
Print #1, "NO. NO.    NAME                       VALUE         DEDN.        SCHEME              "
Print #1, "-------------------------------------------------------------------------------------"
SL = 1
While Not temp.EOF
ps = 0
bval = 0
bval = temp!pbasic + temp!pvda
ps = Round(((temp!pbasic + temp!pvda) * 0.0833), 0)
Print #1, SL & Space(4 - Len(SL)) & Space(4 - Len(temp!T_NO)) & temp!T_NO & Space(3) & Left(temp!expr1, 24) & Space(25 - Len(Left(temp!expr1, 24))) _
& Space(9 - Len(FormatNumber(bval))) & FormatNumber(bval) _
& Space(11 - Len(FormatNumber(temp!epf))) & FormatNumber(temp!epf) _
& Space(16 - Len(FormatNumber(ps))) & FormatNumber(ps) & Space(11 - Len(FormatNumber(temp!epf - ps))) & FormatNumber(temp!epf - ps)

sBVAL = sBVAL + bval
sPS = sPS + ps
sEPF = sEPF + temp!epf
sFPF = sFPF + (temp!epf - ps)
SL = SL + 1
temp.MoveNext

Wend
Print #1, "-------------------------------------------------------------------------------------"
Print #1, Space(45 - Len(FormatNumber(sBVAL))) & FormatNumber(sBVAL) & Space(11 - Len(FormatNumber(sEPF))) & FormatNumber(sEPF) _
& Space(16 - Len(FormatNumber(sPS))) & FormatNumber(sPS) & Space(11 - Len(FormatNumber(sFPF))) & FormatNumber(sFPF)
Print #1, "-------------------------------------------------------------------------------------"
Close #1
RichTextBox1.FileName = "c:\ranjan.txt"
End Sub

Sub printesirepo()
Dim bval, sESI As Double

Open "C:\RANJAN.TXT" For Output As #1
countrec = 0
HEAD = "STATEMENT OF ESI DEDUCTION FOR THE MONTH OF " & MonthName(Month(temp!Month)) & " " & Year(temp!Month)
Print #1, Chr(27) & Chr(67) & Chr(0) & Chr(12)
Print #1,
Print #1, Space((84 - Len("SUMIT ENTERPRISES.")) / 4) & Chr(27) & "W1" & "SUMIT ENTERPRISES." & Chr(27) & "W0"
Print #1, Space((84 - Len(HEAD)) / 2) & UCase(HEAD)
Print #1,
Print #1, "--------------------------------------------------------------------------------"
Print #1, "SL NO.   TICKET     EMPLOYEE NAME                    BASIC            ESI"
Print #1, "         NUMBER                                      VALUE         DEDUCTION"
Print #1, "--------------------------------------------------------------------------------"
SL = 1
While Not temp.EOF
ps = 0
bval = 0
bval = temp!pbasic + temp!pvda
Print #1, SL & Space(4 - Len(SL)) & Space(10 - Len(temp!T_NO)) & temp!T_NO & Space(6) & Left(temp!expr1, 24) & Space(25 - Len(Left(temp!expr1, 24))) _
& Space(14 - Len(FormatNumber(bval))) & FormatNumber(bval) & Space(14 - Len(FormatNumber(temp!esi))) & FormatNumber(temp!esi)
sBVAL = sBVAL + bval
sESI = sESI + temp!esi
SL = SL + 1
temp.MoveNext
Wend
Print #1, "-------------------------------------------------------------------------------------"
Print #1, Space(59 - Len(FormatNumber(sBVAL))) & FormatNumber(sBVAL) & Space(14 - Len(FormatNumber(sESI))) & FormatNumber(sESI)
Print #1, "-------------------------------------------------------------------------------------"
Close #1
RichTextBox1.FileName = "c:\ranjan.txt"

End Sub

Sub PRINTADVANCE()
Dim sADVANCE As Double

Open "C:\RANJAN.TXT" For Output As #1
countrec = 0
HEAD = "STATEMENT OF ADVANCE DEDUCTION FOR THE MONTH OF " & MonthName(Month(temp!Month)) & " " & Year(temp!Month)
Print #1, Chr(27) & Chr(67) & Chr(0) & Chr(12)
Print #1,
Print #1, Space((84 - Len("SUMIT ENTERPRISES.")) / 4) & Chr(27) & "W1" & "SUMIT ENTERPRISES." & Chr(27) & "W0"
Print #1, Space((84 - Len(HEAD)) / 2) & UCase(HEAD)
Print #1,
Print #1, "--------------------------------------------------------------------------------"
Print #1, "SL NO.   TICKET     EMPLOYEE NAME                                ADVANCE"
Print #1, "         NUMBER                                                 DEDUCTION"
Print #1, "--------------------------------------------------------------------------------"
SL = 1
While Not temp.EOF
If Not temp!advrec = 0 Then
sADVANCE = sADVANCE + temp!advrec

Print #1, SL & Space(4 - Len(SL)) & Space(10 - Len(temp!T_NO)) & temp!T_NO & Space(6) & Left(temp!expr1, 24) & Space(25 - Len(Left(temp!expr1, 24))) _
& Space(26 - Len(FormatNumber(temp!advrec))) & FormatNumber(temp!advrec)
SL = SL + 1
End If

temp.MoveNext
Wend
Print #1, "-------------------------------------------------------------------------------------"
Print #1, Space(71 - Len(FormatNumber(sADVANCE))) & FormatNumber(sADVANCE)
Print #1, "-------------------------------------------------------------------------------------"

Close #1
RichTextBox1.FileName = "c:\ranjan.txt"

End Sub

Sub printlic()
Dim sADVANCE As Double

Open "C:\RANJAN.TXT" For Output As #1
countrec = 0
HEAD = "STATEMENT OF LIC DEDUCTION FOR THE MONTH OF " & MonthName(Month(temp!Month)) & " " & Year(temp!Month)
Print #1, Chr(27) & Chr(67) & Chr(0) & Chr(12)
Print #1,
Print #1, Space((84 - Len("SUMIT ENTERPRISES.")) / 4) & Chr(27) & "W1" & "SUMIT ENTERPRISES." & Chr(27) & "W0"
Print #1, Space((84 - Len(HEAD)) / 2) & UCase(HEAD)
Print #1,
Print #1, "--------------------------------------------------------------------------------"
Print #1, "SL NO.   TICKET     EMPLOYEE NAME                                  LIC"
Print #1, "         NUMBER                                                 DEDUCTION"
Print #1, "--------------------------------------------------------------------------------"
SL = 1
While Not temp.EOF
If Not temp!LIC = 0 Then
sADVANCE = sADVANCE + temp!LIC

Print #1, SL & Space(4 - Len(SL)) & Space(10 - Len(temp!T_NO)) & temp!T_NO & Space(6) & Left(temp!expr1, 24) & Space(25 - Len(Left(temp!expr1, 24))) _
& Space(26 - Len(FormatNumber(temp!LIC))) & FormatNumber(temp!LIC)
SL = SL + 1
End If

temp.MoveNext
Wend
Print #1, "-------------------------------------------------------------------------------------"
Print #1, Space(71 - Len(FormatNumber(sADVANCE))) & FormatNumber(sADVANCE)
Print #1, "-------------------------------------------------------------------------------------"

Close #1
RichTextBox1.FileName = "c:\ranjan.txt"

End Sub

Sub printsummary()
Dim MANP As Double
'Dim P As String
'P = Format(temp!Month, "DD-MMM-YYYY")
If temp.State = adStateOpen Then temp.Close
temp.Open "SELECT SUM(ADVREC) AS EXPR1, SUM (PBASIC + PVDA + SPALL + ARRPAY + OTAMT + HRA + CONV) " _
     & "AS TOTER,SUM(EPF + ADVREC + LOAN + LIC + ESI + OTDED) AS TOTDD, " _
    & "SUM (PBASIC + PVDA + SPALL + ARRPAY + OTAMT + HRA + CONV) " _
    & " - SUM(EPF + ADVREC + LOAN + LIC + ESI + OTDED) AS NET, " _
    & "SUM(OTAMT) AS EXPR2, SUM(ARRPAY) AS EXPR3, SUM(HRA)" _
    & "AS EXPR4, SUM(CONV) AS EXPR5, SUM(SPALL) AS EXPR6, " _
    & "SUM(ESI) AS EXPR7, SUM(EPF) AS EXPR8, SUM(ADVREC) " _
    & "AS EXPR9, SUM(LOAN) AS EXPR10, SUM(LIC) AS EXPR11, " _
    & "COUNT(*) AS EXPR12, SUM(OTHRS) AS EXPR13,sum(day_wk) as expr15,sum(day_pay) as expr16, " _
    & "SUM(CL + PL) As EXPR14 From ATTD WHERE MONTH = '" & p & "' ", cn, adOpenKeyset, adLockOptimistic
MANP = (temp!EXPR15 / (temp!EXPR12 * DAYFIND(p)) * 100)
Open "C:\RANJAN.TXT" For Output As #1
countrec = 0
HEAD = "SUMMARY OF PAYMENT FOR THE MONTH OF " & MonthName(Month(p)) & " " & Year(p)
Print #1, Chr(27) & Chr(67) & Chr(0) & Chr(12)
Print #1,
Print #1, Space((84 - Len("SUMIT ENTERPRISES.")) / 4) & Chr(27) & "W1" & "SUMIT ENTERPRISES." & Chr(27) & "W0"
Print #1, Space((84 - Len(HEAD)) / 2) & UCase(HEAD)
Print #1,
Print #1, "SALARY DETAILS:"
Print #1,
Print #1, "GROSS SALARY                : Rs." & Space(16 - Len(FormatNumber(temp!toter))) & FormatNumber(temp!toter)
Print #1, "NET SALARY                  : Rs." & Space(16 - Len(FormatNumber(temp!net))) & FormatNumber(temp!net)
Print #1, "TOTAL OVERTIME PAYMENT      : Rs." & Space(16 - Len(FormatNumber(temp!expr2))) & FormatNumber(temp!expr2)
Print #1, "TOTAL ARREARS PAYMENT       : Rs." & Space(16 - Len(FormatNumber(temp!expr3))) & FormatNumber(temp!expr3)
Print #1, "TOTAL H.R.A PAYMENT         : Rs." & Space(16 - Len(FormatNumber(temp!expr4))) & FormatNumber(temp!expr4)
Print #1, "TOTAL CONVEYANCE ALLOWANCE  : Rs." & Space(16 - Len(FormatNumber(temp!expr5))) & FormatNumber(temp!expr5)
Print #1, "TOTAL SPECIAL ALLOWANCE     : Rs." & Space(16 - Len(FormatNumber(temp!expr6))) & FormatNumber(temp!expr6)
Print #1, "TOTAL ESIC DEDUCTION        : Rs." & Space(16 - Len(FormatNumber(temp!expr7))) & FormatNumber(temp!expr7)
Print #1, "TOTAL PF DEDUCTION          : Rs." & Space(16 - Len(FormatNumber(temp!expr8))) & FormatNumber(temp!expr8)
Print #1, "ADVANCE RECOVERED           : Rs." & Space(16 - Len(FormatNumber(temp!expr9))) & FormatNumber(temp!expr9)
Print #1, "LOAN RECOVERED              : Rs." & Space(16 - Len(FormatNumber(temp!expr10))) & FormatNumber(temp!expr10)
Print #1, "INTEREST ON LOAN            : "
Print #1, "L.I.C. DEDUCTION            : Rs." & Space(16 - Len(FormatNumber(temp!expr11))) & FormatNumber(temp!expr11)
Print #1, "UNION DEDUCTION             : Rs."
Print #1, "---------------------------------------------------------------------------"
Print #1, "ATTENDENCE DETAILS:"
Print #1,
Print #1, "NUMBER OF EMPLOYEES         :  " & Space(18 - Len(FormatNumber(temp!EXPR12))) & FormatNumber(temp!EXPR12)
Print #1, "GROSS MANDAYS AVAILABLE     :  " & Space(18 - Len(FormatNumber(temp!expr16))) & FormatNumber(temp!expr16)
Print #1, "TOTAL OVERTIME HOURS        :  " & Space(18 - Len(FormatNumber(temp!expr13))) & FormatNumber(temp!expr13)
Print #1, "TOTAL LEAVE TAKEN           :  " & Space(18 - Len(FormatNumber(temp!expr14))) & FormatNumber(temp!expr14)
Print #1, "NET MANDAYS AVAILABLE       :  " & Space(18 - Len(FormatNumber(temp!EXPR15))) & FormatNumber(temp!EXPR15)
Print #1, "PERCENTAGE ATTENDENCE       :  " & Space(18 - Len(FormatNumber(MANP))) & FormatNumber(MANP)
Print #1, "---------------------------------------------------------------------------"
Close #1
RichTextBox1.FileName = "c:\ranjan.txt"
End Sub

Sub printcheck()
Dim nBASIC As Double
Dim nVDA, nOTAMT, nHRA, nCONV, nSPALL As Double
Dim nARRPAY, nGROSS, nEPF, nesi, nLOAN As Double
Dim nADV_REC, nLIC, nTOTAL, nNET As Double

Open "C:\RANJAN.TXT" For Output As #1
countrec = 0
SL = 1
HEAD = "ATTENDANCE CHECKPRINT FOR THE MONTH OF " & MonthName(Month(temp!Month)) & " " & Year(temp!Month)
Print #1, Chr(27) & Chr(67) & Chr(0) & Chr(12)
Print #1,
Print #1, Space((84 - Len("SUMIT ENTERPRISES.")) / 2) & Chr(27) & "W1" & "SUMIT ENTERPRISES." & Chr(27) & "W0"

Print #1, Space((84 - Len(HEAD))) & UCase(HEAD) & Chr(15)
'Print #1, String(78, "-")
Print #1, "-----------------------------------------------------------------------------------------------------------------------------------"
Print #1, "SL.  EMP.  NAME                   DAYS        BALANCES            TAKEN            OT   SPL.   ARREARS ADVANCE   LOAN   L.I.C"
 Print #1, "NO.  NO.                     --------------   ----------   ----------------------- HRS  ALLOW          RECOVERY  DED.   DEDNS"
Print #1, "                               WORKED  PAID     CL    EL     CL   EL   ML  ABS LWOP"
Print #1, "-----------------------------------------------------------------------------------------------------------------------------------"
While Not temp.EOF
Print #1, SL & Space(5 - Len(SL)) & Left(temp!expr1, 24) & Space(25 - Len(Left(temp!expr1, 24))) & Space(8 - Len(FormatNumber(temp!BASIC))) & FormatNumber(temp!BASIC) & Space(10 - Len(FormatNumber(temp!day_wk))) & FormatNumber(temp!day_wk) _
& Space(7 - Len(FormatNumber(temp!day_pay))) & FormatNumber(temp!day_pay) & Space(6 - Len(FormatNumber(temp!CL))) & FormatNumber(temp!CL) & Space(6 - Len(FormatNumber(temp!PL))) & FormatNumber(temp!PL) & Space(6 - Len(FormatNumber(temp!lwp))) & FormatNumber(temp!lwp) _
& Space(11 - Len(FormatNumber(temp!pbasic))) & FormatNumber(temp!pbasic) & Space(10 - Len(FormatNumber(temp!pvda))) & FormatNumber(temp!pvda) & Space(9 - Len(FormatNumber(temp!otamt))) & FormatNumber(temp!otamt) & Space(8 - Len(FormatNumber(temp!HRA))) & FormatNumber(temp!HRA) _
& Space(10 - Len(FormatNumber(temp!conv))) & FormatNumber(temp!conv) & Space(10 - Len(FormatNumber(temp!SPALL))) & FormatNumber(temp!SPALL) & Space(8 - Len(FormatNumber(temp!arrpay))) & FormatNumber(temp!arrpay) & Space(12 - Len(FormatNumber(temp!toter))) & FormatNumber(temp!toter) _
& Space(8 - Len(FormatNumber(temp!epf))) & FormatNumber(temp!epf) & Space(8 - Len(FormatNumber(temp!esi))) & FormatNumber(temp!esi) & Space(9 - Len(FormatNumber(temp!loan))) & FormatNumber(temp!loan) & Space(11 - Len(FormatNumber(temp!advrec))) & FormatNumber(temp!advrec) _
& Space(8 - Len(FormatNumber(temp!LIC))) & FormatNumber(temp!LIC) & Space(11 - Len(FormatNumber(temp!TOTDD))) & FormatNumber(temp!TOTDD) _
& Space(11 - Len(FormatNumber(temp!toter - temp!TOTDD))) & FormatNumber(temp!toter - temp!TOTDD)

Print #1, Space(11) & temp!T_NO
Print #1, Space(5) & temp!exprd
Print #1, "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"

'Print #1, 30 - Len(Left(temp!expr1, 24))

'If countrec >= 2 Then
'countrec = 0
'Print #1, Chr(12)
'End If
nBASIC = nBASIC + temp!pbasic
nVDA = nVDA + temp!pvda
nOTAMT = nOTAMT + temp!otamt
nHRA = nHRA + temp!HRA
nCONV = nCONV + temp!conv
nSPALL = nSPALL + temp!SPALL
nARRPAY = nARRPAY + temp!arrpay
nGROSS = nGROSS + temp!toter
nEPF = nEPF + temp!epf
nesi = nesi + temp!esi
nLOAN = nLOAN + temp!loan
nADV_REC = nADV_REC + temp!advrec
nLIC = nLIC + temp!LIC
nTOTAL = nTOTAL + temp!TOTDD
nNET = nNET + (nGROSS - nTOTAL)
SL = SL + 1

countrec = countrec + 1

temp.MoveNext

Wend
Print #1, Space(72) & Space(12 - Len(FormatNumber(nBASIC))) & FormatNumber(nBASIC) & Space(10 - Len(FormatNumber(nVDA))) & FormatNumber(nVDA) _
& Space(9 - Len(FormatNumber(nOTAMT))) & FormatNumber(nOTAMT) & Space(8 - Len(FormatNumber(nHRA))) & FormatNumber(nHRA) _
& Space(10 - Len(FormatNumber(nCONV))) & FormatNumber(nCONV) & Space(10 - Len(FormatNumber(nSPALL))) & FormatNumber(nSPALL) _
& Space(8 - Len(FormatNumber(nARRPAY))) & FormatNumber(nARRPAY) & Space(12 - Len(FormatNumber(nGROSS))) & FormatNumber(nGROSS) _
& Space(8 - Len(FormatNumber(nEPF))) & FormatNumber(nEPF) & Space(8 - Len(FormatNumber(nesi))) & FormatNumber(nesi) _
& Space(9 - Len(FormatNumber(nLOAN))) & FormatNumber(nLOAN) & Space(11 - Len(FormatNumber(nADV_REC))) & FormatNumber(nADV_REC) _
& Space(8 - Len(FormatNumber(nLIC))) & FormatNumber(nLIC) & Space(11 - Len(FormatNumber(nTOTAL))) & FormatNumber(nTOTAL) _
& Space(11 - Len(FormatNumber(nNET))) & FormatNumber(nNET)
Print #1, "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"

Close #1


RichTextBox1.FileName = "c:\ranjan.txt"


End Sub
