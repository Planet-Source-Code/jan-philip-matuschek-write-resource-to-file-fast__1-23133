<div align="center">

## Write Resource To File \- Fast\!


</div>

### Description

This is a fast way to write a resource to a file. UPDATE: You can find an example to this code <a href="http://www.planet-source-code.com/xq/ASP/txtCodeId.23155/lngWId.1/qx/vb/scripts/ShowCode.htm">here</a>.
 
### More Info
 
Filename - Path to the output file, ResID - ID of the resource to write, ResType - Type of the resource to write, Overwrite - Overwrite the output file if it already exists


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jan Philip Matuschek](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jan-philip-matuschek.md)
**Level**          |Beginner
**User Rating**    |3.3 (13 globes from 4 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jan-philip-matuschek-write-resource-to-file-fast__1-23133/archive/master.zip)





### Source Code

```
Public Sub ResToFile(Filename As String, ResID As Variant, ResType As Variant, Optional Overwrite As Boolean = False)
Dim Buffer() As Byte
Dim Filenum As Integer
If Dir(Filename) <> Empty Then 'Check if output file already exists
 If Overwrite Then Kill Filename Else Err.Raise 58
End If
Buffer = LoadResData(ResID, ResType) 'Load the resource into a byte array
Filenum = FreeFile
Open Filename For Binary Access Write As Filenum
Put Filenum, , Buffer 'Write the entire array into the file
Close Filenum
End Sub
```

