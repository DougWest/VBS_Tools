'Option Explicit

Dim wshShell, nBuildNumber, fso, fAssemblyInfo, ReadAllFile, sLine, index, index2

If WScript.Arguments.Count < 1 Then
  WScript.Quit 10
ElseIf WScript.Arguments.Count < 2 Then
  Set wshShell = CreateObject( "WScript.Shell" )
  nBuildNumber = wshShell.ExpandEnvironmentStrings( "%BUILD_NUMBER%" )
ElseIf IsNumeric(WScript.Arguments.Item(1)) Then
  nBuildNumber = WScript.Arguments.Item(1)
Else
  WScript.Quit 11
End If

Set fso = CreateObject("Scripting.FileSystemObject")

If Not fso.FileExists(WScript.Arguments.Item(0)) Then WScript.Quit 12

Set fAssemblyInfo = fso.OpenTextFile(WScript.Arguments.Item(0), 1)
If fAssemblyInfo.AtEndOfStream Then
  ReadAllFile = ""
Else
  ReadAllFile = fAssemblyInfo.ReadAll
End If
fAssemblyInfo.Close
ReadAllFile = Split(ReadAllFile, VBCRLF)

Set fAssemblyInfo = fso.OpenTextFile(WScript.Arguments.Item(0), 2)
For Each sLine in ReadAllFile
  If InStr(sLine, "AssemblyFileVersion") <> 0 Then
    index = InStr(sLine, "(")
    index = InStr(index, sLine, ".")
    index = InStr(index+1, sLine, ".")
    index2 = InStr(index+1, sLine, ".")
    sLine = Left(sLine, index) + nBuildNumber + Mid(sLine, index2 )
  End If
  fAssemblyInfo.Write sLine & VbCRLF
Next
fAssemblyInfo.Close