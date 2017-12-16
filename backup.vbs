'Sebijks Adressverwaltung 1.07
'Copyright 2005 - 2010, Home of the Sebijk.com. Alle Rechte vorbehalten.
'http://www.sebijk.com
'Copyright-Hinweis nicht entfernen.

quellpfad = ".\"
version = "1.07"
Set WshShell = CreateObject("Wscript.Shell")
Set Fs	= CreateObject("Scripting.FileSystemObject")
Set AppShell = CreateObject("Shell.Application")
Set umgebung=WshShell.Environment("PROCESS")
windir=umgebung("windir")
temp=umgebung("temp")
mydoc = wshshell.SpecialFolders("MyDocuments")
Set BrowsePfad = Appshell.BrowseForFolder(0, "Bitte wählen Sie ein Ordner aus, wo die Adressverwaltung gesichert werden sollen.",  &H0001, 17)
On Error Resume Next
If BrowsePfad = "" Then wscript.quit
Ordner = BrowsePfad.ParentFolder.ParseName(BrowsePfad.Title).Path
	If err.number > 0 Then
	i=instr(BrowsePfad, ":")
	Ordner = mid(BrowsePfad, i - 1, 1) & ":\"
	End If
	
If not (Fs.FolderExists(Ordner)) Then
	g = msgbox("Wählen Sie bitte einen gültigen Ordner aus.",vbCritical,"Sebijks Adressverwaltung")
	wscript.quit
End If
i=wshshell.popup("Sie haben als Zielpfad " & Ordner & " eingegeben! Die Adress-Sicherung wird unter " & Ordner & "adressen gespeichert." _
& vbCr & " Wollen Sie fortfahren?",,"Sebijks Adressverwaltung",4 + vbQuestion)
if i = 6 then 
if not fs.FolderExists(Ordner) then fs.CreateFolder(Ordner)
	Set wbasjBatch = fs.OpenTextFile( temp & "\Backup.bat", 2, True)
		wbasjbatch.Writeline "@echo off"
		wbasjbatch.Writeline "echo Sebijks Adressverwaltung " & version
		wbasjbatch.Writeline "echo Copyright 2005 - 2010, Home of the Sebijk.com. Alle Rechte vorbehalten."
		wbasjbatch.Writeline "echo http://www.sebijk.com"
		wbasjbatch.Writeline "REM Copyright-Hinweis nicht entfernen."
		wbasjbatch.Writeline "if ""%OS%""==""Windows_NT"" title Sebijks Adressverwaltung " & version
		wbasjbatch.Writeline "echo."
		wbasjbatch.Writeline "md " & Ordner & "\adressen"
		wbasjbatch.Writeline "copy " & quellpfad & " " & Ordner & "\adressen"
		wbasjbatch.Writeline "copy " & mydoc & "\adressen.mdb " & Ordner & "\adressen"
		wbasjbatch.Writeline "echo  This is a Installation with Global Profile Store. DO NOT DELETE THIS FILE. > " & Ordner & "\adressen\globalprofile.txt"
		wbasjbatch.Writeline "del " & Ordner & "\adressen\unins*"
		wbasjbatch.Writeline "del " & Ordner & "\adressen\backup.*"
		wbasjbatch.Writeline "pause"
		wbasjbatch.Close
		WshShell.Run temp & "\Backup.bat",,true
		fs.DeleteFile temp & "\Backup.bat"
		MsgBox "Die Adressen wurden erfolgreich gesichert!", 64, "Sebijks Adressverwaltung"
End If