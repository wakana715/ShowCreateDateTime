' m4aファイルの「メディアの作成日時」を表示する
Option Explicit
Dim objArgs
Dim strPath
Dim objFso, strFolder, strFile
Dim objShell, objFolder, objFile
Dim txtDateTime
Dim arySlash, arySpace, aryCoron
Dim nYear, nMonth, nDay, nHour, nMinute
Dim strDateTime
Set objArgs = WScript.Arguments
If objArgs.Count = 0 Then
	WScript.Echo "m4aファイルの「メディアの作成日時」を表示する"
	WScript.Echo "CScript /nologo ShowCreateDateTime.vbs <フルパス>"
	WScript.Quit(1)
End If
strPath = objArgs(0)
Set objFso = CreateObject("Scripting.FileSystemObject")
strFolder = objFso.GetParentFolderName(strPath)
strFile = objFso.GetFileName(strPath)
Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.NameSpace(strFolder)
Set objFile = objFolder.ParseName(strFile)
'208 メディアの作成日時
txtDateTime = objFolder.GetDetailsOf(objFile, 208)
arySlash = Split(txtDateTime, "/")
nYear = Int(Mid(arySlash(0), 2))
nMonth = Int(Mid(arySlash(1), 2))
arySpace = Split(arySlash(2), " ")
nDay = Int(Mid(arySpace(0), 2))
aryCoron = Split(arySpace(1), ":")
nHour = Int(Mid(aryCoron(0), 3))
nMinute = Int(aryCoron(1))
strDateTime = _
	Right("000" & CStr(nYear), 4) & _
	Right("0" & CStr(nMonth), 2) & _
	Right("0" & CStr(nDay), 2) & "-" & _
	Right("0" & CStr(nHour), 2) & _
	Right("0" & CStr(nMinute), 2)
WScript.Echo "ren " & Chr(34) & strFile & Chr(34) & _
	" " & Chr(34) & strDateTime + "_" + strFile & Chr(34)
