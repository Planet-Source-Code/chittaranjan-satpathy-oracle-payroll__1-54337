VERSION 5.00
Begin VB.Form FRMSPLASH 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   3225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4485
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   0
      Picture         =   "FRMSPLASH.frx":0000
      ScaleHeight     =   4215
      ScaleWidth      =   4575
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BackStyle       =   0  'Transparent
         Caption         =   "Payroll 1.0"
         BeginProperty Font 
            Name            =   "Haettenschweiler"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   2
         Left            =   840
         TabIndex        =   3
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BackStyle       =   0  'Transparent
         Caption         =   "Chromatics Software"
         BeginProperty Font 
            Name            =   "Haettenschweiler"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   1680
         Width           =   2655
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BackStyle       =   0  'Transparent
         Caption         =   "Connecting Oracle..."
         BeginProperty Font 
            Name            =   "Haettenschweiler"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   2400
         Width           =   2295
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   5400
      Top             =   720
   End
End
Attribute VB_Name = "FRMSPLASH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim temp As New ADODB.Recordset

Private Sub Timer1_Timer()
OPENdatabase
If temp.State = adStateOpen Then temp.Close
temp.Open "select * from para", cn, adOpenKeyset, adLockOptimistic
vda = temp(0)
esi = temp(1)
epf = temp(2)
esicut = temp(3)
'frmLogin.Show
Load MDIForm1
MDIForm1.Show

Unload Me
End Sub
