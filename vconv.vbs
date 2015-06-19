Option Explicit

'' ----------------------------------------------------------------------
''
'' 関数宣言
''
'' ----------------------------------------------------------------------
Function p(s)
    WScript.StdOut.WriteLine s
End Function

Function DirFormat(d)
    DirFormat = Mid(d, 1, 4) & Mid(d, 6, 2) & Mid(d, 9, 2)
End Function

Function NameFormat(d,n)
    NameFormat = DirFormat(d) & "_" & Mid(d, 12, 2) & Mid(d, 15, 2) & Mid(d, 18, 2) & "_" & n & ".mp4"
End Function

'' ----------------------------------------------------------------------
''
'' 変数宣言
''
'' ----------------------------------------------------------------------
Dim reg

Dim objFileSystem
Dim objFolder
Dim objFile
Dim objWShell

Dim PROGRAM
Dim OP_STRING
Dim FileCount

' スクリプトが実行されているパスを取得する
Dim strScriptPath

strScriptPath = Replace(WScript.ScriptFullName, WScript.ScriptName, vbNullString)

PROGRAM = "ffmpeg"
OP_STRING = "-f mp4 -ab 192k -ar 44100 -ac 2 -s 1280x720 -vb 2048k"

Set objWShell = CreateObject("WScript.Shell")

Set reg = CreateObject("VBScript.RegExp")

Set objFileSystem = CreateObject("Scripting.FileSystemObject")

Set objFolder = objFileSystem.GetFolder(strScriptPath)

'' ----------------------------------------------------------------------
''
'' メイン処理
''
'' ----------------------------------------------------------------------
reg.pattern = "\.m.*ts$"
reg.IgnoreCase = True
FileCount = 0
For Each objFile In objFolder.Files
    ' MTSファイルを対象にする
    If reg.Test(objFile.Name) Then
        FileCount = FileCount + 1
        Dim d
        Dim exppath
        Dim expfile

        d = FormatDateTime(objFile.DateCreated, 2) & " " & FormatDateTime(objFile.DateCreated, 4)
        p objFile.Name

        exppath = strScriptPath & "\" & DirFormat(d) 
        expfile = exppath & "\" & NameFormat(d, FileCount)

        If objFileSystem.FolderExists(exppath) = False Then
            objFileSystem.CreateFolder exppath
        end if

        objWShell.Run PROGRAM & " -i """ & objFile.Name & """ " & OP_STRING & " """ & expfile & """", 1, true
    End If
Next

Set objFolder = nothing

Set objFileSystem = Nothing

Set reg = Nothing

set objWShell = Nothing

