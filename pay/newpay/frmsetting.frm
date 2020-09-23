VERSION 5.00
Begin VB.Form frmsetting 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parameter Setting"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   3810
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   10
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   9
      Top             =   720
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2535
      Begin VB.TextBox txtESICUTOFF 
         DataField       =   " "
         DataMember      =   " "
         DataSource      =   "DataEnvironment1"
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
         Left            =   1200
         TabIndex        =   8
         Top             =   960
         Width           =   1000
      End
      Begin VB.TextBox txtEPF 
         DataField       =   " "
         DataMember      =   " "
         DataSource      =   "DataEnvironment1"
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
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1000
      End
      Begin VB.TextBox txtESI 
         DataField       =   " "
         DataMember      =   " "
         DataSource      =   "DataEnvironment1"
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
         Left            =   1200
         TabIndex        =   4
         Top             =   360
         Width           =   1000
      End
      Begin VB.TextBox txtVDA 
         DataField       =   " "
         DataMember      =   " "
         DataSource      =   "DataEnvironment1"
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
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1000
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "ESICUTOFF:"
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
         Left            =   1200
         TabIndex        =   7
         Top             =   720
         Width           =   1170
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "EPF:"
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
         TabIndex        =   5
         Top             =   720
         Width           =   435
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "ESI:"
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
         Left            =   1560
         TabIndex        =   3
         Top             =   120
         Width           =   360
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "VDA:"
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
         Left            =   360
         TabIndex        =   1
         Top             =   120
         Width           =   465
      End
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000004&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   2  'Horizontal Line
      Height          =   165
      Left            =   0
      Top             =   1560
      Width           =   4335
   End
End
Attribute VB_Name = "frmsetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim temp As New ADODB.Recordset

Private Sub Command1_Click()
If Command1.Caption = "Update" Then
Command1.Enabled = False
cn.Execute "update para set vda = " & txtVDA & ",esi= " & txtESI & ",epf = " & txtEPF & ",esicutoff = " & txtESICUTOFF & ""

If temp.State = adStateOpen Then temp.Close
temp.Open "select * from para", cn, adOpenKeyset, adLockOptimistic
vda = temp(0)
esi = temp(1)
epf = temp(2)
esicut = temp(3)
End If

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
If temp.State = adStateOpen Then temp.Close
temp.Open "select * from para", cn, adOpenKeyset, adLockPessimistic
txtEPF = temp!epf
txtESI = temp!esi
txtESICUTOFF = temp!esicutoff
txtVDA = temp!vda
End Sub
