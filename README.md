<div align="center">

## ConvertToSoundex


</div>

### Description

Converts a name or word string to a four digit code following Soundex rules.

Similar code is used by geniological groups and the US Census Bureau for

looking up names by phonetic sound. For example, the name Darrell can

be spelled many different ways. Regardles of how you spell it, (Daryl, Derrel,

Darel, etc.) the Soundex code is always D640. Therefore, you assign a field

in your database to the Soundex code and then query the database using

the code, all instances of Darrell regarless of spelling will be returned. Refer

to the code comment section for more information.
 
### More Info
 
A single name or word string.

A four digit alphanumeric Soundex code.

This code has not been commercially tested.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Darrell Sparti, MCSD](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/darrell-sparti-mcsd.md)
**Level**          |Unknown
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/darrell-sparti-mcsd-converttosoundex__1-1061/archive/master.zip)





### Source Code

```
'***********************************************************************
'Function Name:  ConvertToSoundex
'Argument:      A single name or word string
'Return value:    A 4 character code based on Soundex rules
'Author:        Darrell Sparti
'EMail:        dsparti@allwest.net
'Date:         9-20-98
'Description:    All Soundex codes have 4 alphanumeric
'             characters, no more and no less, regardless
'             of the length of the string. The first
'             character is a letter and the other 3 are
'             numbers. The first letter of the string is
'             the first letter of the Soundex code. The
'             3 digits are defined sequentially from the
'             string using the following key:
'               1 = bpfv
'               2 = cskgjqxz
'               3 = dt
'               4 = l
'               5 = mn
'               6 = r
'               No Code = aehiouyw
'             If the end of the string is reached before
'             filling in 3 numbers, 0's complete the code.
'             Example: Swartz  = S632
'             Example: Darrell  = D640
'             Example: Schultz = S432
'NOTE:        I have noticed some errors in other versions
'            of soundex code. Most noticably is the
'            fact that not only must the code ignore
'            the second letter in repeating letters
'            (ll,rr,tt,etc. for example), it must also
'            ignore letters next to one another with the
'            same soundex code (s and c for example).
'            Other wise, in the example above, Schultz
'            would return a value of S243 which is
'            incorrect.
'********************************************************************
Option Explicit
Public Function ConvertToSoundex(sInString As String) As String
  Dim sSoundexCode As String
  Dim sCurrentCharacter As String
  Dim sPreviousCharacter As String
  Dim iCharacterCount As Integer
  'Convert the string to upper case letters and remove spaces
  sInString = UCase$(Trim(sInString))
  'The soundex code will start with the first character _
  of the string
  sSoundexCode = Left(sInString, 1)
  'Check the other characters starting at the second character
  iCharacterCount = 2
  'Continue the conversion until the soundex code is 4 _
  characters long regarless of the length of the string
  Do While Not Len(sSoundexCode) = 4
   'If the previous character has the same soundex code as _
   current character or the previous character is the same _
   as the current character, ignor it and move onto the next
   sCurrentCharacter = Mid$(sInString, iCharacterCount, 1)
   sPreviousCharacter = Mid$(sInString, iCharacterCount - 1, 1)
   If sCurrentCharacter = sPreviousCharacter Then
     iCharacterCount = iCharacterCount + 1
   ElseIf InStr("BFPV", sCurrentCharacter) Then
     If InStr("BFPV", sPreviousCharacter) Then
      iCharacterCount = iCharacterCount + 1
     End If
   ElseIf InStr("CGJKQSXZ", sCurrentCharacter) Then
     If InStr("CGJKQSXZ", sPreviousCharacter) Then
      iCharacterCount = iCharacterCount + 1
     End If
   ElseIf InStr("DT", sCurrentCharacter) Then
      If InStr("DT", sPreviousCharacter) Then
        iCharacterCount = iCharacterCount + 1
      End If
   ElseIf InStr("MN", sCurrentCharacter) Then
      If InStr("MN", sPreviousCharacter) Then
        iCharacterCount = iCharacterCount + 1
      End If
   Else
   End If
   'If the end of the string is reached before there are 4 _
   characters in the soundex code, add 0 until there are _
   a total of 4 characters in the code
   If iCharacterCount > Len(sInString) Then
     sSoundexCode = sSoundexCode & "0"
   'Otherwise, concatenate a number to the soundex code _
   base on soundex rules
   Else
     sCurrentCharacter = Mid$(sInString, iCharacterCount, 1)
     If InStr("BFPV", sCurrentCharacter) Then
      sSoundexCode = sSoundexCode & "1"
     ElseIf InStr("CGJKQSXZ", sCurrentCharacter) Then
      sSoundexCode = sSoundexCode & "2"
     ElseIf InStr("DT", sCurrentCharacter) Then
      sSoundexCode = sSoundexCode & "3"
     ElseIf InStr("L", sCurrentCharacter) Then
      sSoundexCode = sSoundexCode & "4"
     ElseIf InStr("MN", sCurrentCharacter) Then
      sSoundexCode = sSoundexCode & "5"
     ElseIf InStr("R", sCurrentCharacter) Then
      sSoundexCode = sSoundexCode & "6"
     Else
     End If
   End If
   'Check the next letter
   iCharacterCount = iCharacterCount + 1
  Loop
  'Return the soundex code for the string
  ConvertToSoundex = sSoundexCode
End Function
```

