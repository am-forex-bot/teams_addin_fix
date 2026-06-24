' Run-Hidden.vbs - runs the command line passed to it with NO visible window.
' Usage:  wscript.exe Run-Hidden.vbs <program> <arg1> <arg2> ...
' Any argument containing a space is re-quoted so paths survive intact.
Option Explicit
Dim shell, cmd, i, a
cmd = ""
For i = 0 To WScript.Arguments.Count - 1
  a = WScript.Arguments(i)
  If InStr(a, " ") > 0 Then
    cmd = cmd & " """ & a & """"
  Else
    cmd = cmd & " " & a
  End If
Next
Set shell = CreateObject("WScript.Shell")
shell.Run Trim(cmd), 0, False   ' 0 = hidden window, False = don't wait
