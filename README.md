<div align="center">

## \_InStrLike


</div>

### Description

This is a combination of InStr and the Like operator. It returns the position of a mask within a string. The parameters are all user friendly variants just like the regular InStr function.

Example:

InStrLike("Test String 123abc45 Stuff","###*##")

returns 13, because 123abc45 matches the mask and it starts at character 13. Hope this is useful to somebody.
 
### More Info
 
Start=Position to start searching, optional

String1=String to search

String2=Mask to search for

intCompareMethod=vbCompareMethod to use, optional

This currently does not account for searching for the literal mask characters, which are normally enclosed in brackets in the mask. If that doesn't make sense to you then you are probably not doing it anyways so don't worry about it.

Returns a variant, which is null if String1 or String2 is null, otherwise returns the position of the mask (String2) within the string (String1). 0 if the mask is not present.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Atul Brad Buono](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/atul-brad-buono.md)
**Level**          |Intermediate
**User Rating**    |4.9 (39 globes from 8 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/atul-brad-buono-instrlike__1-11139/archive/master.zip)





### Source Code

```
Public Function InStrLike(Optional ByVal Start, Optional ByVal String1, Optional ByVal String2, Optional ByVal intCompareMethod As VbCompareMethod = vbTextCompare) As Variant
On Error GoTo err_InStrLike
 Dim intPos As Integer
 Dim intLength As Integer
 Dim strBuffer As String
 Dim blnFound As Boolean
 Dim varReturn As Variant
 If Not IsNumeric(Start) And IsMissing(String2) Then
 String2 = String1
 String1 = Start
 Start = 1
 End If
 If IsNull(String1) Or IsNull(String2) Then
 varReturn = Null
 GoTo exit_InStrLike
 End If
 If Left(String2, 1) = "*" Then
 err.Raise vbObjectError + 2600, "InStrLike", "Comparison mask cannot start with '*' since a start position cannot be determined."
 Exit Function
 End If
 For intPos = Start To Len(String1) - Len(String2) + 1
 If InStr(1, String2, "*", vbTextCompare) Then
  For intLength = 1 To Len(String1) - intPos + 1
  strBuffer = Mid(String1, intPos, intLength)
  If strBuffer Like String2 Then
   blnFound = True
   GoTo done
  End If
  Next intLength
 Else
  strBuffer = Mid(String1, intPos, Len(String2))
  If strBuffer Like String2 Then
  blnFound = True
  GoTo done
  End If
 End If
 Next intPos
done:
 If blnFound = False Then
 varReturn = 0
 Else
 varReturn = intPos
 End If
exit_InStrLike:
 InStrLike = varReturn
 Exit Function
err_InStrLike:
 Select Case err.Number
 Case Else
  varReturn = Null
  MsgBox err.Description, vbCritical, "Error #" & err.Number & " (InStrLike)"
  GoTo exit_InStrLike
 End Select
End Function
```

