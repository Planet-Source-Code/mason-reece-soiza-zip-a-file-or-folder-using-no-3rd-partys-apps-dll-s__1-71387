<div align="center">

## Zip a file or folder using no 3rd Partys apps/DLL'S


</div>

### Description

This code will zip any file or folder
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Mason Reece Soiza](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mason-reece-soiza.md)
**Level**          |Intermediate
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mason-reece-soiza-zip-a-file-or-folder-using-no-3rd-partys-apps-dll-s__1-71387/archive/master.zip)





### Source Code

```
Private Sub Command1_Click()
  CreateEmptyZip "c:\testzip.zip"
  With CreateObject("Shell.Application")
    .NameSpace("c:\testzip.zip").CopyHere "c:\FirePassword.exe"
    ' .NameSpace("c:\testzip.zip").CopyHere .NameSpace(FolderName).items 'use this line if we want to zip all items in a folder into our zip file
  End With
  ' All done!
End Sub
Public Sub CreateEmptyZip(sPath)
  Dim strZIPHeader As String
  strZIPHeader = Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0) ' header required to convince Windows shell that this is really a zip file
  With CreateObject("Scripting.FileSystemObject")
    .CreateTextFile(sPath).Write strZIPHeader
  End With
End Sub
```

