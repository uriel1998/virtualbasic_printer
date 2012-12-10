' ##############################################################################
'
' Print Return Transport Request Via Network Printer and Virtual Basic Script
'
' by Steven Saus
'
' No warranty express or implied.
'
' This is a VBS script to print a return transport request with user input
' The printer name is currently hardcoded in.  Requires WinXP or higher,
' and a VBS interpreter.  Obviously, you'll want to have permissions to print
' to the printer.  Printing is handled using the system "print" command so 
' that we can easily deal with non-default printers, network printers, etc.
'
' This is the public version of the script - I have stripped all path names and
' institution identifying information from it.  You will need to put in your own
' printer path names and so on.  These are annotated in TOCHANGE.TXT in this Git
' repository.
'
' Please note that this script explicitly overwrites its temporary files and 
' closes out objects in memory to preserve HIPPA-required privacy.
'
' I had to research a lot of very common VBS problems to compile this script,
' so I'm hoping that putting this script up will give people many examples of 
' those problems.
'
' In particular:
' * Creating the equivalent of a radio button (sort of) from InputBox
' * Comparing strings in VBS
' * Printing to the non-default printer (including networked printers) from CLI
' * Creating multi-line output in a MsgBox
' * Handling case and multi-step if/then loops in VBS
' * Testing for empty strings in VBS
' * Testing for numeric input in VBS
' * Reading and writing from text files in VBS
' * Get a return status from a called system process
' * Execute different actions based on the return status
' * Getting human-readable system time from VBS
' 
'  Licensed under a Creative Commons BY-SA 3.0 Unported license
'  To view a copy of this license, visit http://creativecommons.org/licenses/by-sa/3.0/.
'
' ##############################################################################


' ######################################################## Initialize Filesystem
Dim objFSO 'As FileSystemObject
Set objFSO = CreateObject("Scripting.FileSystemObject")

Dim strDirectory 'As String
strDirectory = "C:\temp"
Dim objDirectory 'As Object

' Ensure that directory exists
If objFSO.FolderExists(strDirectory) Then
	Set objDirectory = objFSO.GetFolder(strDirectory)
Else
	Set objDirectory = objFSO.CreateFolder(strDirectory)
End If

Dim strFile 'As String
strFile = "returnfile.txt"
Dim objTextFile 'As Object
Dim blnOverwrite 'As Boolean
blnOverwrite = True

' ########################################################### Opening Text File
Set objTextFile = objFSO.CreateTextFile(strDirectory & "\" & strFile, blnOverwrite)

' ########################################################### Initialize Vars

Dim strName 'As String
Dim strRoom 'As String
Dim strHolding 'As String
Dim strMode 'As String
Dim strIsol 'As String
Dim strPrinter 'As String
Dim strTime 'As String


' scratch vars
Dim bullet
bullet = Chr(10) & "   " & Chr(149) & " "	
Dim newline
newline = Chr(13) & Chr(10)

' ########################################################### Get User Input Loop
Do
	strName = InputBox("Please type the patient's name: Last, First", "Name Entry", "")
	If strName = "" Then WScript.Quit  'Detect Cancel

	' Getting Room Number
	
		response = InputBox("Please enter the room number where the patient is going.", "Room Entry")
		If response = "" Then WScript.Quit  'Detect Cancel
		strRoom = response

	' Getting Holding Room Number
	Do
		response = InputBox("Please enter HOLDING room number where the patient is located.", "Holding Room Entry")
		If response = "" Then WScript.Quit  'Detect Cancel
		If IsNumeric(response) And response < 99 Then Exit Do 'Detect value response.
		MsgBox "You must enter a valid numeric value.", 48, "Invalid Entry"
	Loop
	strHolding = CStr(response)

	' Setting transport type
	Do
		response = InputBox("Please enter the number that corresponds to the transport mode:" & Chr(10) & bullet & "1 - Wheelchair" & bullet & "2 - Cart" & bullet & "3 - Bed" & Chr(10), "Transport Mode")
		If response = "" Then WScript.Quit  'Detect Cancel
		If IsNumeric(response) And response < 4 Then Exit Do 'Detect value response.
		MsgBox "You must enter a valid numeric value.", 48, "Invalid Entry"
	Loop
	If response = "1" Then strMode = "Wheelchair" 
	If response = "2" Then strMode = "Cart" 
	If response = "3" Then strMode = "Bed" 

	'Setting Isolation Status
	result = MsgBox ("Is the patient in isolation?", vbYesNo, "Isolation Status")
	Select Case result
	Case vbYes
		strIsol = "Isolation Precautions"
	Case vbNo
		strIsol = "Standard Precautions Only"
	End Select

	' Setting output printer
	Do
		response = InputBox("Please enter the number that corresponds to the desired printer:" & Chr(10) & bullet & "1 - PRINTERONE" & bullet & "2 - PRINTERTWO" & bullet & "3 - PRINTERTHREE" & Chr(10), "Printer Selection")
		If response = "" Then WScript.Quit  'Detect Cancel
		If IsNumeric(response) And response < 4 Then Exit Do 'Detect value response.
		MsgBox "You must enter a valid numeric value.", 48, "Invalid Entry"
	Loop
	If response = "1" Then strPrinter = "\\PATHTO\PRINTER" 
	If response = "2" Then strPrinter = "\\PATHTO\PRINTER" 
	If response = "3" Then strPrinter = "\\PATHTO\PRINTER" 

	' Setting time
	d=Now()
	strTime = FormatDateTime(d,2) & " " & FormatDateTime(d,4)

' ########################################################### Check User Input

	result = MsgBox ("Is the following correct?" &newline & "Patient: " & strName & newline & "Room: " & strRoom & newline & "Holding Room: " & strHolding & newline & "Transport: " & strMode & newline & "Precautions: " & strIsol & newline & "Request time: " & strTime & newline & "Desired printer: " & strPrinter,vbYesNo,"Verification")
		If result = vbYes Then Exit Do
Loop

' WScript.Quit


' ########################################################### Write the Text File
objTextFile.WriteLine("################################################################################")
objTextFile.WriteLine("                   RETURN TRANSPORT REQUEST")
objTextFile.WriteLine( newline & newline & newline & "     Patient:      " & strName & newline & "     Patient Room: " & strRoom & newline & "     Holding Room: " & strHolding & newline & "     Transport:    " & strMode & newline & "     Precautions:  " & strIsol & newline & "     Request time: " & strTime & newline & newline & newline)
objTextFile.WriteLine(newline & newline & newline)
objTextFile.WriteLine("################################################################################")

' Closing everything
objTextFile.Close

' ################################################################### Print File

'Printing file 
Set WshShell = WScript.CreateObject("WScript.Shell")
Return = WshShell.Run("%comspec% /c print /d:" & strPrinter & " c:\temp\returnfile.txt > c:\temp\returnstatus.txt", 0, true)

' ############################################################### Error Checking

' Need to open the return status file and compare it to a known output for error catching.
Set objReturnFile = objFSO.OpenTextFile("c:\temp\returnstatus.txt", 1)
text = objReturnFile.ReadAll
objReturnFile.Close

Dim strPrintError 'As String

strPrintError = "Unable to initialize device " & strPrinter
return = StrComp(text,strPrintError,1) 
' This option could be used to create a logfile if needed....
If return = "-1" Then MsgBox "The request is currently being printed."
If return = "1" Then MsgBox " " & newline & "##############################################" & newline & "There has been an error printing to " & strPrinter &"." & newline & "Please contact transport via phone." & newline & "##############################################"

' ################################################################## Closing out
' Overwriting temp file
Set objTextFile = objFSO.CreateTextFile(strDirectory & "\" & strFile, blnOverwrite)
objTextFile.Close

Set objReturnFile = Nothing
Set objTextFile = Nothing
Set objDirectory = Nothing
Set objFSO = Nothing
