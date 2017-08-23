'Set options on the script.
Option Explicit

'Declare all variables upfront.
Dim objWMIService, objProcess, colProcess, strComputer, processName, processEndName, instances, maxInstances, oShell, objFSO, outFile, objFile, timestamp

'The computer to apply this to (. means "this one")
strComputer = "."

'Placeholder to count how many processes there are.
instances = 0

'The max number of the process that can run before
'the program kills them.
maxInstances = 10

'The name of the process to count.
processName = "P21CrystalIntegration.exe"

'This is the process that will be killed.
processEndName = "splwow64.exe"

'Name of the file to write to.
outFile = "log.txt"

'Execution code below here onward.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Set objWMIService = GetObject("winmgmts:" _
					& "{impersonationLevel=impersonate}!\\" _
					& strComputer & "\root\cimv2")

Set colProcess = objWMIService.ExecQuery _
				 ("SELECT * FROM Win32_Process")

For Each objProcess in colProcess
	If objProcess.Name = processName Then
		instances = instances + 1
	End If
Next

Set objFSO = CreateObject("Scripting.FileSystemObject")

'If the file exists, open it for appending (code 8).
'Otherwise, create the file and then write to it.
If objFSO.FileExists(outFile) Then
	Set objFile = objFSO.OpenTextFile(outFile, 8, True)
Else
	Set objFile = objFSO.CreateTextFile(outFile, True)
End If

Set oShell = CreateObject("WScript.Shell")

timestamp = Now()

objFile.Write("[" & timestamp & "] Found Instances: " & instances & "/" & maxInstances & " of " & processName & vbCrlf)
If instances >= maxInstances Then
	objFile.Write("[" & timestamp & "] Killing " & processEndName & "..." & vbCrlf)
	oShell.Run "taskkill /im " & processEndName, , True
	objFile.Write("[" & timestamp & "] Process terminated." & vbCrlf)
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

objFile.Write(" " & vbCrlf)
objFile.Close