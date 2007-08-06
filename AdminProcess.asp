<!--#INCLUDE FILE="adovbs.inc"-->
<%

'*** 5:10 AM 3/7/2007 ~ MZD ***
' * "Fixing" Page Name functionality - auto-file-stub generation
' * really, fully implementing for the first time
' * noticing that Session("new") may be always on and not necessary to be checked.  
' *Seems to have been replaced by Addit, which is a overloaded state variable
'*** 

function CreateNewFileStub(PageName, Path, FSO)
	filedef = FSO.BuildPath(Path, PageName & ".asp")
	NewPageName = PageName
	n = 1
	Do While FSO.FileExists(filedef)
		NewPageName = PageName & n 
		filedef = FSO.BuildPath(Path, NewPageName & ".asp" )
		n = n + 1
	loop

	call objFSO.CopyFile("d:\websites\sligosite\Admin\template.asp", filedef)

	CreateNewFileStub = NewPageName
End Function
'~~~

'	IF Not IsEmpty(Request.form("id")) THEN 
		id = Request.Form("id")	
		menuid = Request.Form("mid")
'DEBUG:
'Response.write("Form ID[" &Request.Form("id") & "]")
'Response.write("Form MID[" &Request.Form("mid") & "]")
'Response.write("Form ADDIT[" &Request.Form("ADDIT") & "]")

'	else
'		id = 0
'	end if
       	
	If Request.Form("title") = "" or IsEmpty(Request.Form("title")) then
		title = "Temp title"
	else
		title = Request.Form("title")
	end if
        tophead = Request.Form("tophead")
        bottomhead = Request.Form("bottomhead")
        graphic = Request.Form("graphic")
        category = Request.Form("category")
        catsort  = Request.Form("catsort")
        sort = Request.Form("sort")
        noshow = Request.Form("noshow")
        description = Request.Form("description")

	IF Session("new") then
		topdescription = Request.Form("topdescription")
		PageName = Request.Form("PageName")
		MenuName = Request.Form("MenuName")
	END IF

' * build sql statement to retrieve all fields

' * Determine if we are going to add or edit
Addit = Request.Form("Addit")

IF (Addit <> "Y") THEN
   sql = "SELECT * FROM Main WHERE ID = " & id & " ;"
ELSE
   sql = "SELECT * FROM Main where id = 0;"
END IF

' * create connection objects
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.connectionstring = Session("connstring")
Conn.Open

set rs = Server.CreateObject("ADODB.Recordset")
RS.Open sql, Conn, adOpenKeyset, adLockOptimistic
	
IF Addit = "D" THEN
	RS.Delete
  	RS.Close
   	set RS = nothing
   	set Conn = nothing
	Response.Redirect ("AdminEntry.asp?mid=" & menuid & "")
END IF

IF (Addit = "Y") THEN 
	RS.AddNew
	RS("MenuID") = MenuID
END IF
	
' * Give (some) record set fields their values
RS("title")    = title
RS("tophead") = tophead
RS("bottomhead") = bottomhead
RS("graphic") = graphic
RS("category") = category
RS("catsort")  = catsort

if sort = "" then 
	sort = 0
end if 

RS("sort") = sort

if noshow = "on" then
	RS("noshow") = -1
else
	rs("noshow") = 0
end if
RS("description") = description

'IF Session("new") then

'MZD: why are these two assignments here??:
	RS("topdescription") = topdescription
	RS("MenuName") = MenuName

	dim objFSO
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

	' create new template file
	If IsNull(RS("menuName")) Then
		newapathdir = ""
	Else
		newpathdir = RS("menuname")
	End If
	NewPath = objFSO.BuildPath(objFSO.BuildPath("d:\websites\sligosite\", Session("path") ), newpathdir)
	
	IF (Addit = "Y") THEN 

		PageName = CreateNewFileStub(PageName, NewPath, objFSO)

	Else 'Addit <> "Y"
		If NOT IsNull(RS("PageName") ) Then
			IF PageName <> RS("PageName") Then
				' * user input does not match db record. Create new file stub and delete old.
				' * Set new PageName value for updating db below
	
				oldFileName = RS("PageName") & ".asp"
				If objFSO.FileExists(objFSO.BuildPath(NewPath, oldFileName)) Then
					call objFSO.DeleteFile(objFSO.BuildPath(NewPath, oldFileName ), True)
				End If
	
				PageName = CreateNewFileStub(PageName, NewPath, objFSO)

			End If
		End If
	END IF

	RS("PageName") = PageName	' for update (below)

	set objFSO = nothing
'END IF

RS.Update ' * Write values to record

Conn.Close

' * free memory
set RS = nothing
set Conn = nothing

Response.Redirect ("AdminEntry.asp?mid=" & menuid & "")
%>
