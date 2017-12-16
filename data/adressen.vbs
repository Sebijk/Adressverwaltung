window.resizeTo 800, 500
version = "1.07"
appname = "Sebijk's Adressverwaltung"
Set wshshell = CreateObject("WScript.Shell")
Set FileSystem = CreateObject("Scripting.FileSystemObject")
Dim scriptname
Dim thisdir
Dim mydoc
Dim sjdir
Dim dbfile
scriptname = URLDecode(document.location.pathname)
thisdir = FileSystem.getparentfoldername(scriptname)
mydoc = wshshell.SpecialFolders("MyDocuments")
sjdir = mydoc

if FileSystem.FileExists(thisdir & "\globalprofile.txt") then sjdir = thisdir
dbfile = FileSystem.buildpath(sjdir, "adressen.mdb")
provider = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=""" & dbfile & """"
' Pruefen, ob adressen.mdb existiert
If not FileSystem.FileExists(dbfile) then 
  Set adox = CreateObject("ADOX.Catalog")
  adox.Create provider
  Set table = CreateObject("ADOX.Table")
  table.Name = "Adressen"
  ' AutoNumber-Feld definieren
  Set column = CreateObject("ADOX.Column")
  column.Name = "id"
  column.Type = 3
  Set column.ParentCatalog = adox
  column.properties("AutoIncrement") = True
  table.Columns.Append column
  ' normales Feld definieren
  table.Columns.Append "name",202,255
  table.Columns("name").Attributes = 3
  ' normales Feld definieren
  table.Columns.Append "vorname",202,255
  table.Columns("vorname").Attributes = 3
  ' normales Feld definieren
  table.Columns.Append "email",202,255
  table.Columns("email").Attributes = 3
  ' normales Feld definieren
  table.Columns.Append "telefon",202,255
  table.Columns("telefon").Attributes = 3
  ' normales Feld definieren
  table.Columns.Append "strasse",202,255
  table.Columns("strasse").Attributes = 3
  ' normales Feld definieren
  table.Columns.Append "hausnr",202,255
  table.Columns("hausnr").Attributes = 3
  ' normales Feld definieren
  table.Columns.Append "plz",202,255
  table.Columns("plz").Attributes = 3
  ' normales Feld definieren
  table.Columns.Append "ort",202,255
  table.Columns("ort").Attributes = 3
  ' normales Feld definieren
  table.Columns.Append "handy",202,255
  table.Columns("handy").Attributes = 3
  ' normales Feld definieren
  table.Columns.Append "homepage",202,255
  table.Columns("homepage").Attributes = 3
  adox.Tables.Append table
End If

function URLDecode(str)
	dim re
	set re = new RegExp

	str = Replace(str, "+", " ")
	
	re.Pattern = "%([0-9a-fA-F]{2})"
	re.Global = True
	URLDecode = re.Replace(str, GetRef("URLDecodeHex"))
end function

Function list
  Set adox = CreateObject("ADOX.Catalog")
  Set oConn = CreateObject("ADODB.Connection")
	oConn.Open provider

	dim strSQL
	strSQL = "SELECT vorname,name,strasse,hausnr,plz,ort,telefon,email,handy,homepage from [adressen] order by name;"

	Dim oRs
	Set oRs = oConn.Execute( strSQL)

	Do while (Not oRs.eof) 
	    InsertRow oRs
		oRs.MoveNext 
	Loop 
End Function

' Replacement function for the above
function URLDecodeHex(match, hex_digits, pos, source)
	URLDecodeHex = chr("&H" & hex_digits)
end function

' auflisten aller eintraege aus der datenbank
function InsertCellHTML( cell, celltext)
   dim str 
   str = "<font face=verdana size=1>" + celltext + "</font>"
   cell.innerHTML = str
end function

function InsertCell( record, oRow, cellname)
	dim text
	dim oCell
	text = record.Fields( cellname)
	set oCell = oRow.insertCell()
	InsertCellHTML oCell, text
end function

function InsertRow( record)
	dim oTable
	set oTable = document.all("adressen")
	dim oRow
	set oRow = oTable.insertRow()

	InsertCell record, oRow, "vorname"
	InsertCell record, oRow, "name"
	InsertCell record, oRow, "strasse"
	InsertCell record, oRow, "hausnr"
	InsertCell record, oRow, "plz"
	InsertCell record, oRow, "ort"
	InsertCell record, oRow, "telefon"
	InsertCell record, oRow, "email"
	InsertCell record, oRow, "handy"
	InsertCell record, oRow, "homepage"
end function

Sub add
  Set oConn = CreateObject("ADODB.Connection")
  oConn.Open provider

Do
	vorname = InputBox("Geben Sie den Vornamen ein.",appname)
	If IsEmpty(vorname) then exit Do
	vorname = Replace(vorname, "'", "''")
	name = InputBox("Geben Sie den Nachnamen ein.",appname)
	If IsEmpty(name) then exit Do
	name = Replace(name, "'", "''")
	strasse = InputBox("Geben Sie die Straße ein.",appname)
	If IsEmpty(strasse) then exit Do
	strasse = Replace(strasse, "'", "''")
	hausnr = InputBox("Geben Sie die Hausnummer ein.",appname)
	If IsEmpty(hausnr) then exit Do
	hausnr = Replace(hausnr, "'", "''")
	plz = InputBox("Geben Sie die PLZ ein.",appname)
	If IsEmpty(plz) then exit Do
	plz = Replace(plz, "'", "''")
	ort = InputBox("Geben Sie den Ort ein.",appname)
	If IsEmpty(ort) then exit Do
	ort = Replace(ort, "'", "''")
	telefon = InputBox("Geben Sie die Telefonnummer ein.",appname)
	If IsEmpty(telefon) then exit Do
	telefon = Replace(telefon, "'", "''")
	email = InputBox("Geben Sie die E-Mail-Adresse ein.",appname)
	If IsEmpty(email) then exit Do
	email = Replace(email, "'", "''")
	handy = InputBox("Geben Sie die Handynummer ein.",appname)
	If IsEmpty(handy) then exit Do
	handy = Replace(handy, "'", "''")
	homepage = InputBox("Geben Sie die Homepage ein.",appname)
	If IsEmpty(homepage) then exit Do
	homepage = Replace(homepage, "'", "''")
	sql = "insert into Adressen (vorname, name, strasse, hausnr, plz, ort, telefon, email, handy, homepage) VALUES ('"
	sql = sql & vorname & "','" & name & "','" & strasse & "','" & hausnr & "','" & plz & "','" & ort & "','" & telefon & "','" & email & "','" & handy & "','" & homepage & "')"
	MsgBox "Sie haben folgendes eingegeben: " & vbCr & "Name: " & vorname & " " & name & vbCr & "Straße und Hausnr.: " & strasse & " " & hausnr & vbCr & "PLZ und Ort: " & plz & " " & ort & vbCr &"Telefon: " & telefon & vbCr & "Handy: " & handy & vbCr &"E-Mail Adresse: " & email & vbCr & "Homepage: " & homepage & vbCr & "Klicken Sie auf OK um fortzufahren.",64,appname
	oConn.Execute sql
	MsgBox "Die Adresse wurde erfolgreich gespeichert.",64,appname
Loop
location.reload()
End Sub

Sub search
Set oConn = CreateObject("ADODB.Connection")
oConn.Open provider
suchname = InputBox("Geben Sie einen Teil des gesuchten Nachnamens ein.",appname)
If IsEmpty(suchname) then exit sub
	
sql = "select * from Adressen where name like '%" & suchname & "%'"

filename = FileSystem.BuildPath(sjdir, "dbresult.htm")
Set file = FileSystem.CreateTextFile(filename, true)
file.WriteLine "<html><head><title>Sebijks Adressverwaltung - Ergebnisse</title>"
file.WriteLine "<!--"
file.WriteLine ""
file.WriteLine appname
file.WriteLine "Version " & version
file.WriteLine "Copyright © 2005 - 2010 Home of the Sebijk.com"
file.WriteLine "Alle Rechte vorbehalten."
file.WriteLine ""
file.WriteLine "!-->"
file.WriteLine "<meta name=""vs_targetSchema"" content=""http://www.sebijk.com"">"
file.WriteLine "<meta http-equiv=""Page-Enter"" content=""blendTrans(Duration=1)"">"
file.WriteLine "<meta http-equiv=""Page-Exit"" content=""blendTrans(Duration=1)"">"
file.WriteLine "<style type=""text/css"">"
file.WriteLine "body {	scrollbar-arrow-color: #000000;"
file.WriteLine "font-family:Verdana;"
file.WriteLine "font-style:normal;"
file.WriteLine "font-size:12;"
file.WriteLine "background-color:Buttonface }"
file.WriteLine "</style>"
file.WriteLine "</head>"
file.WriteLine "<body dir=""LTR"" background=""background.jpg"" bgproperties=""fixed"">"
file.WriteLine "<font face=""Arial,Helvetica,Verdana"" size=""5"">Sebijks Adressverwaltung: Suchergebnisse</font>"
file.WriteLine "<hr size=""1"" />"
file.WriteLine "<table border=""1"">"
file.WriteLine "<tr>"
Set rs = oConn.Execute(sql)
For each field in rs.fields
  file.WriteLine "<th><font face=Tahoma size=1>" & field.name & "</font></th>"
Next
file.WriteLine "</tr>"
Do until rs.eof
  file.WriteLine "<tr>"
  For each field in rs.fields
    file.WriteLine "<td><font face=Tahoma size=1>" & field.value & "</font></td>"
  Next
  file.WriteLine "</tr>"
  rs.MoveNext
Loop

file.WriteLine "</table></body></html>"
file.close
wshshell.Run """" & filename & """"
End Sub

Sub deletequery
Set oConn = CreateObject("ADODB.Connection")
oConn.Open provider
Do
	vorname = InputBox("Geben Sie den Vornamen ein.",appname)
	If IsEmpty(vorname) then exit Do
	vorname = Replace(vorname, "'", "''")
	name = InputBox("Geben Sie den Nachnamen ein.",appname)
	If IsEmpty(name) then exit Do
	name = Replace(name, "'", "''")
	sql = "DELETE FROM Adressen WHERE vorname = '" & vorname & "' AND name = '" & name & "';"
	oConn.Execute sql
	MsgBox "Der Adressdatensatz wurde erfolgreich gelöscht.",64,appname
Loop
location.reload()
End Sub


Sub saveastxt
Set oConn = CreateObject("ADODB.Connection")
oConn.Open provider
sql = "select * from Adressen"

filename = FileSystem.BuildPath(sjdir, "dbresult.txt")
Set file = FileSystem.CreateTextFile(filename, true)

Set rs = oConn.Execute(sql)
file.WriteLine "Sebijks Adressverwaltung:"
file.WriteBlankLines(1)
Do until rs.eof
  For each field in rs.fields
    file.WriteLine field.name & ": " & field.value
  Next
  file.WriteBlankLines(1)
  rs.MoveNext
Loop
file.close
wshshell.Run """" & filename & """"
End Sub

Sub del
Set adox = CreateObject("ADOX.Catalog")
kennwort = InputBox("Bitte geben Sie das Kennwort ein.",appname)
If kennwort = "" then MsgBox "Sie haben nichts eingegeben! Die Adressen werden nicht gelöscht!", vbCritical, appname
If kennwort = "" then exit sub
if not kennwort = "twreEkbSL19AX" then
  MsgBox "Das Kennwort ist ungültig! Die Adressen werden nicht gelöscht!", vbCritical, appname
  exit sub
End If
If kennwort = "twreEkbSL19AX" then
antwort = MsgBox("Möchten Sie wirklich alle Adressen löschen?", vbYesNo+vbExclamation+vbSystemModal, appname)
  If antwort = vbNo then 
     exit sub
  else
     FileSystem.DeleteFile dbfile,true
     adox.Create provider
     Set table = CreateObject("ADOX.Table")
     table.Name = "Adressen"
     ' AutoNumber-Feld definieren
     Set column = CreateObject("ADOX.Column")
     column.Name = "id"
     column.Type = 3
     Set column.ParentCatalog = adox
     column.properties("AutoIncrement") = True
     table.Columns.Append column
     ' normales Feld definieren
     table.Columns.Append "name",202,255
     table.Columns("name").Attributes = 3
     ' normales Feld definieren
     table.Columns.Append "vorname",202,255
     table.Columns("vorname").Attributes = 3
     ' normales Feld definieren
     table.Columns.Append "email",202,255
     table.Columns("email").Attributes = 3
     ' normales Feld definieren
     table.Columns.Append "telefon",202,255
     table.Columns("telefon").Attributes = 3
     ' normales Feld definieren
     table.Columns.Append "strasse",202,255
     table.Columns("strasse").Attributes = 3
     ' normales Feld definieren
     table.Columns.Append "hausnr",202,255
     table.Columns("hausnr").Attributes = 3
     ' normales Feld definieren
     table.Columns.Append "plz",202,255
     table.Columns("plz").Attributes = 3
     ' normales Feld definieren
     table.Columns.Append "ort",202,255
     table.Columns("ort").Attributes = 3
     ' normales Feld definieren
     table.Columns.Append "handy",202,255
     table.Columns("handy").Attributes = 3
     ' normales Feld definieren
     table.Columns.Append "homepage",202,255
     table.Columns("homepage").Attributes = 3
     adox.Tables.Append table
     MsgBox "Alle Adressen wurden gelöscht.", 64, appname
     location.reload()
  End If
  else
  End If
End Sub

Sub printer
  ergebnis = MsgBox("Bevor Sie die Adressen drucken, stellen Sie sicher dass Sie die Adressen aufgelistet haben. " _
  & vbCr &  "Es empfiehlt sich außerdem die Adressen im Querformat auszudrucken. Wollen Sie jetzt Drucken?", vbYesNo + vbQuestion, appname)
  if ergebnis = vbNo then exit sub
  window.print()
end sub

Sub reload
  location.reload()
end sub

Sub backup
  quellpfad = ".\"
  Set AppShell = CreateObject("Shell.Application")
  Set umgebung=WshShell.Environment("PROCESS")
  windir=umgebung("windir")
  temp=umgebung("temp")
  Set BrowsePfad = Appshell.BrowseForFolder(0, "Bitte wählen Sie ein Ordner aus, wo die Adressverwaltung gesichert werden soll!",  &H0001, 17)
  On Error Resume Next
  If BrowsePfad = "" Then exit sub
  Ordner = BrowsePfad.ParentFolder.ParseName(BrowsePfad.Title).Path
    If err.number > 0 Then
      i=instr(BrowsePfad, ":")
      Ordner = mid(BrowsePfad, i - 1, 1) & ":\"
    End If
	
  If not (FileSystem.FolderExists(Ordner)) Then
    g = msgbox("Wählen Sie bitte einen gültigen Ordner aus.",vbCritical,appname)
    exit sub
  End If

i=wshshell.popup("Sie haben als Zielpfad " & Ordner & " eingegeben! Die Adressen werden unter " & Ordner & "adressen gespeichert." _
& vbCr & " Wollen Sie fortfahren?",,appname,4 + vbQuestion)
if i = 6 then 
if not FileSystem.FolderExists(Ordner) then FileSystem.CreateFolder(Ordner)
if not FileSystem.FolderExists(Ordner & "\adressen") then FileSystem.CreateFolder(Ordner & "\adressen")
FileSystem.CopyFile dbfile , Ordner & "\adressen\adressen.mdb", OverwriteExisting
MsgBox "Die Adressen wurden erfolgreich gesichert.", 64, appname
End If
End Sub

Sub restore
kennwort = InputBox("Bitte geben Sie das Kennwort ein.",appname)
If kennwort = "" then MsgBox "Sie haben nichts eingegeben. Die Adressen werden nicht wiederhergestellt.", vbCritical, appname
If kennwort = "" then exit sub
If kennwort = "twreEkbSL19AX" then
quellpfad = ".\"
Set AppShell = CreateObject("Shell.Application")
Set umgebung=WshShell.Environment("PROCESS")
windir=umgebung("windir")
temp=umgebung("temp")
Set BrowsePfad = Appshell.BrowseForFolder(0, "Bitte wählen Sie ein Ordner aus, wo sich die Adressdatenbankdatei befindet.",  &H0001, 17)
On Error Resume Next
If BrowsePfad = "" Then exit sub
Ordner = BrowsePfad.ParentFolder.ParseName(BrowsePfad.Title).Path
	If err.number > 0 Then
	i=instr(BrowsePfad, ":")
	Ordner = mid(BrowsePfad, i - 1, 1) & ":\"
	End If
	
If not (FileSystem.FolderExists(Ordner)) Then
	g = msgbox("Wählen Sie bitte einen gültigen Ordner aus.",vbCritical,appname)
	exit sub
End If

i=wshshell.popup("Sie haben als Pfad " & Ordner & " eingegeben." & vbCr & "" & vbCr _
& "ACHTUNG: Alle Adressen die Sie zuletzt eingegeben hatten, werden überschrieben. Möchten Sie fortfahren?",vbExclamation,appname,4 + vbQuestion)
if i = 6 then 
if not FileSystem.FileExists(Ordner & "\adressen.mdb") then 
zielpfad = Ordner & "\adressen.mdb"
Ordner = zielpfad
end if
if FileSystem.FileExists(Ordner & "\adressen.mdb") then 
FileSystem.DeleteFile quellpfad & "adressen.mdb"
FileSystem.CopyFile Ordner & "\adressen.mdb" , quellpfad , OverWriteFiles
MsgBox "Die Adressen wurden erfolgreich wiederhergestellt!", 64, appname
else
MsgBox "Die Datei ""adressen.mdb"" ist in " & zielpfad & " nicht vorhanden.", vbCritical, appname
End If
  else
  if not kennwort = "twreEkbSL19AX" then
  MsgBox "Das Kennwort ist ungültig! Die Adressen werden nicht wiederhergestellt!", vbCritical, appname
  end if
  exit sub
  End If
End If
End Sub

Sub Info
  MsgBox appname & vbCr & "Version " & version & vbCr &  "Copyright © 2005 - 2010 Home of the Sebijk.com" & vbCr & "Verwendete Teile: dhtmlxMenu (www.dhtmlx.com)" & vbCr & "" & vbCr & _
  "Kennwort für löschen/wiederherstellen: twreEkbSL19AX" & vbCr & "" & vbCr & "Dieses Programm ist lizenziert unter der GPL.", vbInformation, "Info über Sebijks Adressverwaltung " & version
end Sub

Sub onQuit
  window.close()
End Sub

Sub Help
  WshShell.Run ".\adressen.chm"
End Sub

Sub show_license
  WshShell.Run ".\license.txt"
End Sub