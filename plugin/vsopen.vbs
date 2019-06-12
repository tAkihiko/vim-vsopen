' Copyed from https://teratail.com/questions/115801
On Error Resume Next

Dim file, line, offset
Dim result
if WScript.Arguments.Count <> 3 then
	WScript.echo("usage: jumpToVS.vbs arg1 arg2 arg3")
	WScript.Quit(-1)
end if
file = WScript.Arguments(0)
line = WScript.Arguments(1)
offset = WScript.Arguments(2)

Set objDTE = GetObject(, "VisualStudio.DTE.14.0")
If Err.Number<> 0 Then
	Set objDTE = CreateObject("VisualStudio.DTE.14.0")
	Err.Clear
End If

objDTE.MainWindow.Activate
objDTE.MainWindow.Visible = True
objDTE.UserControl = True

objDTE.ItemOperations.OpenFile(file)
result = objDTE.ActiveDocument.Selection.MoveToLineAndOffset(line, offset)
