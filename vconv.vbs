Option Explicit

Function DirFormat(d)
    DirFormat = Mid(d, 1, 4) & Mid(d, 6, 2) & Mid(d, 9, 2)
End Function

Function NameFormat(d)
    NameFormat = DirFormat(d) & "_" & Mid(d, 12, 2) & Mid(d, 15, 2) & Mid(d, 18, 2) & ".mp4"
End Function

Dim reg
Dim stdout

Dim objFileSystem
Dim objFolder
Dim objFile

Dim strScriptPath

strScriptPath = Replace(WScript.ScriptFullName, WScript.ScriptName, vbNullString)

Set reg = CreateObject("VBScript.RegExp")

Set objFileSystem = CreateObject("Scripting.FileSystemObject")

Set objFolder = objFileSystem.GetFolder(strScriptPath)

reg.pattern = "^..*\.M.*TS$"
For Each objFile In objFolder.Files
    If reg.Test(objFile.Name) Then
        Dim d
        d = FormatDateTime(objFile.DateCreated)
        Wscript.StdOut.WriteLine objFile.Name
        Wscript.StdOut.WriteLine DirFormat(d) & "/" & NameFormat(d)
    End If
Next

Set objFolder = nothing

Set objFileSystem = Nothing


Set reg = Nothing

