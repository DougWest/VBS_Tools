'Source
'Option Explicit

Dim wshShell, nBuildNumber, nRevision, fso, fAssemblyInfo, ReadAllFile, sLine, index, index2, index3

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
nRevision = "0"
If (3 = WScript.Arguments.Count) Then If "SNAPSHOT" = WScript.Arguments.Item(2) Then nRevision = "9999"
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
  If (InStr(WScript.Arguments.Item(0), ".vbp") <> 0) Then 'This logic if the file is a VB6 project file.
    If (InStr(sLine, "RevisionVer") <> 0) Then
	  sLine = "RevisionVer=" + nBuildNumber
	End If
  ElseIf Left(sLine, 2) <> "//" Then 'This logic if the file is a C# AssemblyInfo.cs file.
    If (InStr(sLine, "AssemblyFileVersion") <> 0) OR (InStr(sLine, "AssemblyVersion") <> 0) Then
      index = InStr(sLine, "(")
      index = InStr(index, sLine, ".")
      index = InStr(index+1, sLine, ".")
      index2 = InStr(index+1, sLine, ".")
	  If index2 < 1 Then 
	    index2 = InStr(index+1, sLine, """")
		index3 = index2
	  Else
	    index3 = InStr(index2+1, sLine, """")
	  End If
	  If 0 = nRevision Then
        sLine = Left(sLine, index) + nBuildNumber + Mid(sLine, index2 )
	  Else
	    sLine = Left(sLine, index) + nBuildNumber + "." + nRevision + Mid(sLine, index3 )
	  End If
    End If
  End If
  fAssemblyInfo.Write sLine & VbCRLF
Next
fAssemblyInfo.Close