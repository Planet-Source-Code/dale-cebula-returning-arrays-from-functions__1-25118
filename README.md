<div align="center">

## Returning Arrays From Functions


</div>

### Description

The following code demonstrates how to call a function and return multiple results in an array.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Dale Cebula](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dale-cebula.md)
**Level**          |Beginner
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dale-cebula-returning-arrays-from-functions__1-25118/archive/master.zip)





### Source Code

For example: You could have a function that returns error information which is called like this:
<br>
<br>
Private Sub MySub()
<br>
On Error GoTo err_handler
<br>
'....code here that rasies an error
<br>
err_handler:
<br>
If Err.Number <> 0 Then
<br>
 Dim Tmp() As String
<br>
 Tmp = ErrorHandler
<br>
 MsgBox "Error Description: " & Tmp(0) & " Error Number #:" & Tmp(1) & " Source: " & Tmp(2)
<br>
Erase Tmp
<br>
End If
End Sub
<br>
<br>
<br>
<br>
Public Function ErrorHandler() As String()
<br>
Dim Errors(0 To 2) As String
<br>
 Errors(0) = Err.Description
<br>
 Errors(1) = Err.Number
<br>
 Errors(2) = Err.Source
<br>
 Err.Clear
<br>
 ErrorHandler = Errors
<br>
End Function

