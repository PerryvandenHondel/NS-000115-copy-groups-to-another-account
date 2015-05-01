

Option Explicit

Const	FOR_READING			=	1
Const	FOR_WRITING   		=	2
Const	FOR_APPENDING		=	8



Dim		gobjFso
Dim		gobjShell
Dim		gstrFolderBatch
Dim		sPreviousDate
Dim		nCount
Dim		x


Call ScriptInit()
Call ScriptRun()
Call ScriptDone()
WScript.Quit(0)



Sub ScriptInit()
	Set gobjShell = CreateObject("WScript.Shell")
	Set gobjFso = CreateObject("Scripting.FileSystemObject")

	gstrFolderBatch = GetScriptPath()
End Sub


Sub ScriptRun()
	Dim		strRootDse	' Root DSE
	Dim		strSa		' Source account
	Dim		strDa		' Destionation account
	
	if WScript.Arguments.Named.Count <> 3 then
		Call ScriptUsage()
		WScript.Quit(0)
	end if

	strRootDse = WScript.Arguments.Named("rootdse")
	strSa = WScript.Arguments.Named("sa")
	strDa = WScript.Arguments.Named("da")
	
	WScript.Echo "In domain " & strRootDse & " the groups of " & strSa & " will be copied to " & strDa
	
	Call DsutilsMakeGroupsSameAs(strRootDse, strSa, strDa)
End Sub


Sub ScriptDone()
	Set gobjFso = Nothing
	Set gobjShell = Nothing
End Sub


Sub ScriptUsage()
	WScript.Echo "Usage:"
	WScript.Echo vbTab & WScript.ScriptName & "  /rootdse:<Root DSE> /sa:<user name> /da:<user name>"
	WScript.Echo
	WScript.Echo vbTab & vbTab & "/rootdse:<path.xlsx>          Root DSE for the domain to connect to"
	WScript.Echo vbTab & vbTab & "/sa:<user name>               Source sAMAccountName"
	WScript.Echo vbTab & vbTab & "/da:<user name>               Destination sAMAccountName"
	WScript.Echo 
End Sub


Function GetScriptName
	''
	''	GetScriptName
	''
	''	Returns:
	''		The script name only without the extension
	''
	Dim	sScriptName
	
	sScriptName = WScript.ScriptName
	WScript.Echo sScriptName
	sScriptName = Mid(sScriptName, 1, InStrRev(sScriptName, ".") - 1)
	GetScriptName = sScriptName
End Function '' GetScriptName


'WScript.Echo GetBatchPath()

' For x = 1 To 1000
'	WScript.Echo x & vbTab & GetBatchPath()

' Next

'WScript.Echo DsutilsGetDnFromSam("DC=prod,DC=ns,DC=nl", "ADB.Notexist")


Sub DsutilsMakeGroupsSameAs(ByVal strRoot, ByVal strSourceSam, ByVal strDestSam)
	'
	'	Make the user groups of strSourceSam the same as strDestSam
	'	Use the DSUTILS for these actions

	Dim		strSourceDn
	Dim		strDestDn
	Dim		c
	Dim		strPathTemp
	Dim		l
	Dim		f
	Dim		ts
	Dim		e
	Dim		b
	Dim		p
	
	' Get the DN's of the sam accounts
	strSourceDn = DsutilsGetDnFromSam(strRoot, strSourceSam)
	If InStr(strSourceDn, "CN=") = 0 Then
		WScript.Echo "ERROR: No DN found for " & strSourceSam & " in " & strRoot
		Exit Sub
	End If
	
	strDestDn = DsutilsGetDnFromSam(strRoot, strDestSam)
	If InStr(strDestDn, "CN=") = 0 Then
		WScript.Echo "ERROR: No DN found for " & strDestSam & " in " & strRoot
		Exit Sub
	End If
	
	'WScript.Echo "strSourceDn=" & strSourceDn
	'WScript.Echo "strDestDn=  " & strDestDn
	
	'	Get the group members of the source DN
	' 	dsget user "CN=BEH_Perry.vdHondel,OU=Admin,OU=Beheer,DC=test,DC=ns,DC=nl" -memberof

	'	Get the temp file path
	strPathTemp = GetPathTempFile()
	
	'	Build the command line
	c = "dsget.exe user " & Chr(34) & strSourceDn & Chr(34) & " -memberof >" & strPathTemp
	gobjShell.Run "cmd /c " & c, 0, True
	 
	Set b = New ClassTextFile
	
	p = GetBatchPath()
	WScript.Echo "Write batch to: " & p
	
	WScript.Echo "DsutilsMakeGroupsSameAs(): SRC=" & strSourceSam & " DST=" & strDestSam & " BATCH=" & p
	
	b.SetMode(FOR_WRITING)
	b.SetPath(p)
	b.OpenFile()
	 
	Set f = gobjFso.GetFile(strPathTemp)
	Set ts = f.OpenAsTextStream(FOR_READING)
	Do While ts.AtEndOfStream <> True
		l = ts.ReadLine
		'WScript.Echo "LINE: " & l
		
		If InStr(l, "CN=") > 0 Then
			'	Only process lines with CN= in the line
			'	dsmod group "CN=TYP_BEHEER_ADB,OU=Global Groups,OU=Beheer,DC=test,DC=ns,DC=nl" -addmbr "CN=BEH_Perry.vdHondel,OU=Admin,OU=Beheer,DC=test,DC=ns,DC=nl"

			e = "dsmod.exe group " & l & " -addmbr " & strDestDn
			''b.WriteLineToFile(e)
			WScript.Echo e

			On Error Resume Next
			gobjShell.Run "cmd /c " & e, 6, True
			WScript.Echo Err.Number
		End If
	Loop
	ts.Close
	f.Delete
	
	b.CloseFile()
	
	gobjShell.Run "cmd /c " & p, 6, True
	
	
	Set b = Nothing
	
	Set ts = Nothing
	Set f = Nothing
End Sub



Function GetBatchPath()
	Dim		p
	Dim		f
	Dim		r
	
	f = GenerateFilenameDT()
	p = gstrFolderBatch & f & ".cmd"
	
	If gobjFso.FileExists(p) = True Then
		r = GetBatchPath()
	Else
		r = p
	End If
	
	GetBatchPath = r
End Function ' GetBatchPath



Function DsutilsGetDnFromSam(ByVal strRoot, ByVal strSamAccount)
	'
	'	Get the DN from an sAMAccountName using DSQUERY.EXE
	'
	'	strRoot			"DC=test,DC=ns,DC=nl"
	'	strSamAccount	"Perry.vandenHondel"
	'
	Const	FOR_READING = 1
	
	Dim		r			'	Result
	Dim		c			'	Command Line
	Dim		f			'	File Object
	Dim		ts			' 	TextStream
	Dim		l			' 	Line
	Dim		i
	Dim		x
	Dim		objShell
	Dim		objExec
	Dim		strPath
	
	Set objShell = CreateObject("WScript.Shell")

	c = "dsquery.exe user " & strRoot & " -samid " & strSamAccount 
	'WScript.Echo c
	Set objExec = objShell.Exec(c)
	r = objExec.StdOut.ReadLine
	'WScript.Echo "r=[" & r & "]"
	Set objShell = Nothing
	
	DsutilsGetDnFromSam = r
End Function ' DsutilsGetDnFromSam



Function GetPathTempFile()
	'
	'	Returns a path to a temp file
	'
	'	4 digits for HHMM
	'	12 digits for unique name of hexadecimal number
	'
	'
	Const	TEMP_FOLDER		=	2
	Const	NUM_LOW			=	0
	Const	NUM_HIGH		=	15
	Const	HEX_LEN			=	16
	
	Dim		objFso
	Dim		strTempFolder
	Dim		strFilename
	Dim		i
	Dim		intNumber
	Dim		strPath
	Dim		dtmNow
	Dim		strNow
	
	Set objFso = CreateObject("Scripting.FileSystemObject")

	strTempFolder = objFso.GetSpecialFolder(TEMP_FOLDER)
	
	dtmNow = Now()
	
	strNow = NumberAlign(Hour(dtmNow), 2) & NumberAlign(Minute(dtmNow), 2)
	
	
	For i = 1 to HEX_LEN - 4
		Randomize	
		intNumber = Int((NUM_HIGH - NUM_LOW + 1) * Rnd + NUM_LOW)
		strFilename = strFilename & LCase(Hex(intNumber))
	Next
	
	'
	'	strTempFolder : c:\temp\
	'	strNow        : 2345
	'	strFilename   : a45f6def812
	'	
	strPath = strTempFolder & "\" & strNow & strFilename & ".tmp"
	 
	If objFso.FileExists(strPath) = True Then
		strPath = GetPathTempFile()
	End If
	
	Set objFso = Nothing
	
	GetPathTempFile = strPath
End Function


Function GetScriptPath()
	''
	''	Returns the path where the script is located.
	''
	''	Output:
	''		A string with the path where the script is run from.
	''
	''		drive:\folder\folder\   REMARK: extra \ at the end!!
	''
	''
	Dim sScriptPath
	Dim sScriptName

	sScriptPath = WScript.ScriptFullName
	sScriptName = WScript.ScriptName

	GetScriptPath = Left(sScriptPath, Len(sScriptPath) - Len(sScriptName))
End Function '' GetScriptPath


Function GetFileNameTemp()
	Const	NUM_LOW			=	0
	Const	NUM_HIGH		=	15
	Const	HEX_LEN			=	16
	
	Dim		dtmNow
	Dim		strNow
	Dim		intNumber
	Dim		strFilename
	Dim		i
	
	dtmNow = Now()
	strNow = NumberAlign(Hour(dtmNow), 2) & NumberAlign(Minute(dtmNow), 2)
	For i = 1 to HEX_LEN - 4
		Randomize	
		intNumber = Int((NUM_HIGH - NUM_LOW + 1) * Rnd + NUM_LOW)
		strFilename = strFilename & LCase(Hex(intNumber))
	Next
	
	GetFileNameTemp = strNow & strFilename
End Function



Function NumberAlign(ByVal intNumber, ByVal intLen)
	'
	'	Returns a number aligned with zeros to a defined length
	'
	'	NumberAlign(1234, 6) returns '001234'
	'
	NumberAlign = Right(String(intLen, "0") & intNumber, intLen)
End Function '' NumberAlign


Function GenerateFilenameDT
		''
		''	Generate a new DateTime string with milliseconds
		''
		''	Format: YYYY-MM-DD HH:MM:SS.XXXXXX
		
		Dim	sCurrentDate
	
		sCurrentDate = ProperDateTimeFs("")
		
		'WScript.Echo "sCurrentDate="&sCurrentDate & vbTab & "sPreviousDate="&sPreviousDate
		
		If sCurrentDate = sPreviousDate Then
			' Dates equal, increase counter
			nCount = nCount + 1
		Else
			' Reset the XXX counter
			nCount = 0
		End If
	
		sPreviousDate = sCurrentDate
		GenerateFilenameDT = sCurrentDate & Right("0000" & CStr(nCount), 4)
	End Function '' GenerateDT



Private Function ProperDateTimeFs(ByVal dDateTime)
	'=
	'=	Convert a system formatted date time to a proper format
	'=	Returns the current date time in proper format when no date time
	'=	is specified.
	'=
	'=	15-5-2009 4:51:57  ==>  20090515045157

	If Len(dDateTime) = 0 Then
		dDateTime = Now()
	End If

	ProperDateTimeFs = NumberAlign(Year(dDateTime), 4) & _ 
		NumberAlign(Month(dDateTime), 2) & _
		NumberAlign(Day(dDateTime), 2) & _
		NumberAlign(Hour(dDateTime), 2) & _
		NumberAlign(Minute(dDateTime), 2) & _
		NumberAlign(Second(dDateTime), 2)
End Function '' ProperDateTime
	


Function GetTempFileName
	'//////////////////////////////////////////////////////////////////////////////
	'//
	'//	GetTempFileName
	'//
	'//	Input:
	'//		None
	'//
	'//	Output:
	'//		A string with a temporary file name, e.g. E:\temp\filename.tmp
	'//
	Dim	oTempFolder
	Dim	sTempFile
	
	Const	WINDOWS_FOLDER		=	0
	Const	SYSTEM_FOLDER		=	1
	Const	TEMPORARY_FOLDER	=	2

	Set oTempFolder = goFSO.GetSpecialFolder(TEMPORARY_FOLDER)
	sTempFile = goFSO.GetTempName
	GetTempFileName = oTempFolder & "\" & sTempFile
	Set oTempFolder = Nothing
End Function '' GetTempFileName



Class ClassTextFile
	'
	'	General class to handle file operations for text files .
	'
	'	Parent class for
	'		ClassTextFileTsv
	'		ClassTextFileSplunk
	'
	'	Class Subs and Functions:
	'		Private Sub Class_Initialize				Class initializer sub, set all default values
	'		Private Sub Class_Terminate					Class terminator, releases all variables, etc..
	'		Public Sub SetPath(ByVal strPathNew)		Sets the path to the file.
	'		Public Function GetPath						Returns the path of the file C:\folder\file.ext
	'		Public Sub SetMode							Set the mode of access for the file (READ, WRITE, APPEND)
	'		Public Sub OpenFile							Open the file
	'		Public Sub CloseFile						Closes the file
	'		Public Sub WriteToFile(ByVal strLine)		Write a line to the file
	'		Public Function ReadFromFile()				Read a line from the  file
	'		Public Sub DeleteFile						Delete the file
	'		Function IsEndOfFile						Boolean returns the end of the file reached
	'		Public Function CurrentLine()				Returns the current line number
	'

	Private		objFso	
	Private		objFile
	Private		strPath
	Private		blnIsOpen
	Private		intMode				'	Modus of file activity, READING=1, WRITING=2, APPENDING=8
	Private		intLineCount
	
	Private Sub Class_Initialize
		'
		'	Class initializer, open objects, set default variable values.
		'
		Set objFso = CreateObject("Scripting.FileSystemObject")
		blnIsOpen = False
		intLineCount = 0
	End Sub '' Class_Initialize

	Private Sub Class_Terminate
		'
		'	Class terminator, closes objects
		'
		Call CloseFile()
		
		'	Terminate the object to the text file
		Set objFile = Nothing
		
		'	Terminate the object to the File System Object
		Set objFso = Nothing
	End Sub '' Class Terminate
	
	Public Sub SetPath(ByVal strPathNew)
		'
		'	Sets the path to the file.
		'	Assumes all folders exist before calling this function
		'
		'If objFso.FileExists(strPathNew) = False Then
		'	WScript.Echo "ClassTextFile.SetPath() ERROR: Path is not found!"
		' End If
		strPath = strPathNew
	End Sub
	
	Public Function GetPath
		'
		'	Returns the current path. c:\folder\file.ext
		'
		GetPath = strPath
	End Function
	
	Public Sub SetMode(intModeNew)
		'
		'	Set the mode of file opening
		'		1:	READ
		'		2:	WRITE
		'		8:	APPEND
		'
		intMode = intModeNew
	End Sub
	
	Public Function GetMode()
		'
		'	Return the mode of file opening.
		'		1:	READ
		'		2:	WRITE
		'		8:	APPEND
		'		
		GetMode = intMode
	End Function
	
	Public Sub OpenFile
		'
		'	Open the file strPath
		'
		On Error Resume Next
		Set objFile = objFso.OpenTextFile(strPath, intMode, True)
		If Err.Number = 0 Then
			blnIsOpen = True
		Else
			WScript.Echo("ClassTextFile/OpenFile ERROR: Could not open textfile: " & strPath)
		End If
	End Sub
	
	Public Sub CloseFile()
		'	
		'	Closes the current opened file
		'	
		If blnIsOpen = False Then
			objFile.Close
		End If
	End Sub
	
	Public Sub WriteLineToFile(ByVal strLine)
		'
		'	Write the contents of strLine to the text file.
		'
		If blnIsOpen = True Then
			objFile.WriteLine(strLine)
		Else
			WScript.Echo "ClassTextFile/WriteLineToFile WARNING: Tried to write to a closed file: " & strPath
		End If
	End Sub
	
	Public Function ReadLineFromFile()
		'
		'	Read a line from the text file.
		'	
		'	Returns a string
		'	Returns a empty string when nothing could be read.
		'
		If blnIsOpen = True Then
			'	Increase the line counter +1
			intLineCount = intLineCount + 1
			
			'	Read a line from the text file.
			ReadLineFromFile = objFile.ReadLine
		Else
			ReadLineFromFile = ""
			WScript.Echo "ClassTextFile/ReadLineFromFile WARNING: Tried to read from a closed file: " & strPath
		End If
	End Function
	
	Public Function CurrentLine()
		'
		'	Return the current line number.
		'
		CurrentLine = intLineCount
	End Function
	
	Public Sub DeleteFile()
		'
		'	Delete the file
		'	Close it if it is open.
		'
		If blnIsOpen = True Then
			'	Close the file if it's open
			Call CloseFile()
			If objFso.FileExists(strPath) Then
				'	Delete the file, always!
				Call objFso.DeleteFile(strPath, True)
			End If
		End If
   	End Sub
	
	Function IsEndOfFile()
		'
		'	Return the AtEndOfStreamStatus
		'	
		'	True	End of stream reached
		'	False	No reached yet
		'
		IsEndOfFile = objFile.AtEndOfStream
	End Function

End Class	'	ClassTextFile
