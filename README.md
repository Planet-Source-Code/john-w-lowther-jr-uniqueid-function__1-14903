<div align="center">

## UniqueID Function


</div>

### Description

One-Liner to generate a unique Id. Does NOT require a database. No autonumber fields needed.
 
### More Info
 
Assumes the system clock is never reset to a previous time. Could generate duplicate's if the system clock is reset to a prevoius time, but not likely due to the precision of the function.

UniqueID based on the System Clock


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[John W\. Lowther, Jr\.](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/john-w-lowther-jr.md)
**Level**          |Advanced
**User Rating**    |4.5 (18 globes from 4 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/john-w-lowther-jr-uniqueid-function__1-14903/archive/master.zip)





### Source Code

```
Public Function GetUniqueId() As String
  GetUniqueId = Trim(Str(CDbl(Now) * 10000000000#))
End Function
```

