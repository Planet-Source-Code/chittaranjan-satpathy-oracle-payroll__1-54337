Attribute VB_Name = "modNumToWord"
' NumToWord V2
'
' Ramu

Public Function NumToWord( _
      TheNumber As Variant, _
      Optional Prefix As String = "Rupees", _
      Optional Suffix As String = "Only.", _
      Optional Deci As String = "& Paisa") As String
  Dim strNum As String
  Dim n As Integer
  
  ' this takes care of the rounding off to 2 decimal places as well
  ' try entering something like .345
  strNum = Format(Val(TheNumber), "#00000000000.00")
  
  ' maximum
  ' Rupees Ninety Nine Hundred Ninety Nine Crores Ninety Nine Lacs Ninety
  ' Nine Thousand Nine Hundred Ninety Nine & Paisa Ninety Nine Only.
  '
  ' if you need more, like you make sales in that many rupees,
  ' let me know, i'll come and work for you.
  '
  If Len(strNum) > 14 Then NumToWord = "Number too long to convert.": Exit Function
   
  ' the Zero is taken care here itself
  If Val(Mid(strNum, 1, 11)) = 0 Then NumToWord = "Zero "
  
  ' conditions to take care of Hundred, Crores, Lacs, Thousand, Hundred, Units & Paise
  n = Val(Mid(strNum, 1, 2)):   If n > 0 Then NumToWord = NumToWord & ToWord(n) & " Hundred "
  
  ' incase there are only zeroes after this till 7 places
  If Val(Mid(strNum, 3, 7)) = 0 And Val(Mid(strNum, 1, 2)) > 0 Then NumToWord = NumToWord & "Crore "
  
  n = Val(Mid(strNum, 3, 2)):   If n > 0 Then NumToWord = NumToWord & ToWord(n) & " Crore "
  
  n = Val(Mid(strNum, 5, 2)):   If n > 0 Then NumToWord = NumToWord & ToWord(n) & " Lac "
  
  n = Val(Mid(strNum, 7, 2)):   If n > 0 Then NumToWord = NumToWord & ToWord(n) & " Thousand "
  
  n = Val(Mid(strNum, 9, 1)):   If n > 0 Then NumToWord = NumToWord & ToWord(n) & " Hundred "

  n = Val(Mid(strNum, 10, 2)):  If n > 0 Then NumToWord = NumToWord & ToWord(n)

  ' if there are any after the "."
  n = Val(Mid(strNum, 13, 2)):  If n > 0 Then NumToWord = Trim(NumToWord) & " " & Deci & " " & ToWord(n)

  If Len(NumToWord) > 0 Then NumToWord = Prefix & " " & Trim(NumToWord) & " " & Suffix

  
End Function



' Converts any 1 - 2 digit number to word.
Private Function ToWord(n As Integer) As String
  Dim aUnit() As String
  Dim aTens() As String
  Dim nDigit ' wonder why "As Integer" is missing ?
  
  aUnit = Split("One,Two,Three,Four,Five,Six,Seven,Eight,Nine,Ten," & _
                "Eleven,Twelve,Thirteen,Fourteen,Fifteen,Sixteen,Seventeen,Eighteen,Nineteen", ",")
  
  aTens = Split("Twenty,Thirty,Forty,Fifty,Sixty,Seventy,Eighty,Ninety", ",")
  
  If n > 19 Then
    nDigit = n / 10
    nDigit = Int(nDigit)
    ToWord = aTens(nDigit - 2)
    n = n - (nDigit * 10)
    n = Int(n)
  End If
  
  If n > 0 Then ToWord = ToWord & " " & aUnit(n - 1)
  
End Function
