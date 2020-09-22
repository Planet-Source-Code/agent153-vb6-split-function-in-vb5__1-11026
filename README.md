<div align="center">

## VB6 Split Function in VB5


</div>

### Description

This code duplicates the functionality of VB6's split function.
 
### More Info
 
TheString$: The string to be parsed

Optional Delim$: The dilimeter to parse TheString$ by

Optional Limit: The maximum number of elements to return.

If Delim is ommited or is not found, a single element array is returned.

To have Split return all of the substring, Limit should be set to -1.

Variant containing an array

If your program is migrated to VB6, this code will need to be removed, as it shares the same functionality (and name) of a VB6 intrisync function.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Agent153](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/agent153.md)
**Level**          |Intermediate
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/agent153-vb6-split-function-in-vb5__1-11026/archive/master.zip)





### Source Code

```
Function Split(TheString As String, Optional Delim As String, Optional Limit As Long = -1) As Variant
  'Duplicates the functionality of the vb6 counterpart.
  'Unfortunately, I was unable to include the vbcompare part of the vb6 funtionality.
  'Just use Option Campare at the beggining of this module.
  Dim dynArray() As Variant
  If Len(Delim) > 0 Then
    Dim ArrCt%
    Dim CurPos%
    Dim LenAssigned%
    Dim CurStrLen%
    ArrCt% = 0
    CurPos% = 1
    LenAssigned% = 1
    CurStrLen% = Len(TheString$)
    Do
      ReDim Preserve dynArray(0 To ArrCt%)
      CurStrLen% = (InStr(CurPos%, TheString$, Delim$) - CurPos%)
      If CurStrLen% < 0 Then
        dynArray(ArrCt%) = Right$(TheString$, (Len(TheString$) - (LenAssigned% - 1)))
        Exit Do
      Else
        dynArray(ArrCt%) = Mid$(TheString$, CurPos%, CurStrLen%)
      End If
      LenAssigned% = LenAssigned% + (Len(dynArray(ArrCt%)) + Len(Delim$))
      CurPos% = LenAssigned%
      ArrCt% = ArrCt% + 1
      If Not Limit = -1 Then
        If ArrCt = Limit Then Exit Do
      End If
    Loop
    Split = dynArray
  Else
    'duplicate the functionality more acuratley
    ReDim dynArray(0 To 0)
    dynArray(0) = TheString
    Split = dynArray
  End If
End Function
```

