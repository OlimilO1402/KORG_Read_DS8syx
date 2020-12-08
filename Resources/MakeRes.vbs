With CreateObject("Scripting.FileSystemObject")
	Dim CurDir: CurDir = .GetParentFolderName(WScript.ScriptFullName)
	MsgBox CurDir
    Dim rcExe:  rcExe  = CurDir & "\rc.exe"
    Dim MyRes:  MyRes  = CurDir & "\MyRes.rc" 
    If Not .FileExists(rcExe) Then MsgBox "Couldn't find rc.exe in:" & vbLf & CurDir: WScript.Quit
    If Not .FileExists(MyRes) Then MsgBox "Couldn't find MyRes.rc in:" & vbLf & CurDir: WScript.Quit
    
	CreateObject("Shell.Application").ShellExecute """" & rcExe & """", """" & MyRes & """" & " >>log.txt", "", "", 1
End With