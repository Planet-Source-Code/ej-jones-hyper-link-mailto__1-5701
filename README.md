<div align="center">

## Hyper\-link / mailto


</div>

### Description

Easily lets you add a link from you VB application or open default email program.
 
### More Info
 
textbox

Follow these instructions and you should have no problems.

1. Open new project

2. Place two text boxes on the form

3. Place two command buttons on form

4. Copy and paste code on form

Open specific web site, or default email


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Ej Jones](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ej-jones.md)
**Level**          |Beginner
**User Rating**    |4.7 (42 globes from 9 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ej-jones-hyper-link-mailto__1-5701/archive/master.zip)





### Source Code

```
Private Sub Command1_Click()
Dim x
x = Shell("start.exe " & Text1, 0)
End Sub
Private Sub Command2_Click()
Dim x
x = Shell("start.exe mailto:" & Text1, 0)
End Sub
```

