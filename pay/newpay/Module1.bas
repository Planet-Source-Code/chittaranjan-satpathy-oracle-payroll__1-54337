Attribute VB_Name = "Module1"
Dim temp As New ADODB.Recordset
Public cn As New ADODB.Connection
Public dt As String
Public days As Integer
'Public Const vda = 1755
Public vda As Double
Public esi As Double
Public epf As Double
Public esicut As Double


Public Sub OPENdatabase()
On Error GoTo MSG
If cn.State = adStateOpen Then cn.Close
'if you are accessing oracle from client computer
' you need to change the service name what you have given
' the SID name at the time of installation in stead of ORACOR

' uncomment the below line if you are connecting from
' client computer

'cn.Open "Provider=MSDAORA.1;Password=challan;User ID=challan;data source = oracor;Persist Security Info=True"

' make comment the below line if you are connecting from
' client computer

cn.Open "Provider=MSDAORA.1;Password=challan;User ID=challan;Persist Security Info=True"

Exit Sub

MSG:
MsgBox "Either you have  not  installed   oracle or " & vbCrLf _
    & "not created user as instructed in readme.txt " & vbCrLf _
     & "or you need to erase the Data Source=ORACOR  " & vbCrLf _
& " follow as instructed in OPENDATABASE FUNCTION ", vbCritical

End


End Sub

' this function is used to convert esi round up

Public Function ROUNDUP(ByVal number As Double) As Double
Dim num As String
Dim last As Double, add As Double
num = Format(number, "########.00")
If Val(Right(num, 1)) <= 5 And Val(Right(num, 1)) <> 0 Then
add = (5 - Val(Right(num, 1))) / 100
End If
If Val(Right(num, 1)) > 5 Then
add = (10 - Val(Right(num, 1))) / 100
End If
ROUNDUP = Val(num) + add
End Function

Sub main()
FRMSPLASH.Show
End Sub

' automatically find's the day paid up

Public Function DAYFIND(ByVal MTH As String) As Double
Select Case Month(MTH)
Case 1, 3, 5, 7, 8, 10, 12
days1 = 31
Case 4, 6, 9, 11
days1 = 30
Case 2
If Int(Year(MTH) / 4) * 4 = Year(MTH) Then
days1 = 29
Else
days1 = 28
End If
End Select
DAYFIND = days1
End Function
