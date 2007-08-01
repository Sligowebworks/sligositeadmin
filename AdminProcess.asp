<!--#INCLUDE FILE="adovbs.inc"-->

 <%
    ' IF Not IsEmpty(Request.form("id")) THEN 
	id = Request.Form("id")	
	menuid = Request.Form("mid")
	'response.Write(menuid)
	'else
	'	id = 0
	'end if
       	if Request.Form("title") = "" or IsEmpty(Request.Form("title")) then
			title = "Temp title"
			
		else
			title = Request.Form("title")
		end if
		'response.Write(title)
		tophead = Request.Form("tophead")
        bottomhead = Request.Form("bottomhead")
        graphic = Request.Form("graphic")
        category = Request.Form("category")
        catsort  = Request.Form("catsort")
        sort = Request.Form("sort")
        noshow = Request.Form("noshow")
        description = Request.Form("description")
	     if Session("new") then
		topdescription = Request.Form("topdescription")
		PageName = Request.Form("PageName")
		menuName = Request.Form("MenuName")
	    end if
' On Error Resume Next
' build an sql statement to retrieve all fields

' Determine if we are going to add or edit
Addit = Request.Form("Addit")

if (Addit <> "Y") then
   sql = "SELECT * FROM Main WHERE ID = " & id & " ;"
else
   sql = "SELECT * FROM Main where id = 0;"
end if

'response.write(sql)
'create connection objects
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.connectionstring = Session("connstring")
Conn.Open 	

set rs = Server.CreateObject("ADODB.Recordset")
RS.Open sql, Conn, adOpenKeyset, adLockOptimistic
	
if Addit = "D" then
	RS.Delete
  	RS.Close
  	set RS = nothing
 	set Conn = nothing
	Response.Redirect ("AdminEntry.asp?mid=" & menuid & "")
end if
if (Addit = "Y")  then 
	RS.AddNew
	RS("MenuID") = MenuID
end if
' Give record set fields their values
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
if Session("new") then
    RS("topdescription") = topdescription
	RS("PageName") = PageName
	RS("MenuName") = MenuName
end if
RS.Update  'Write values to record set
' create new template file
if (Addit = "Y") then 
	dim objFSO
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	vfile = "D:\Websites\SligoSite\CDHP\" & Session("path") & "\" & RS("menuname") & "\"  & RS("PageName") & ".asp"
	response.write("d:\websites\sligosite\" & Session("path") & "\" & "template.asp" & "<BR>")
	response.write(vfile)
	call objFSO.CopyFile("D:\Websites\SligoSite\CDHP\" & Session("path") & "\" & "template.asp", vfile)
end if
Conn.Close
' free memory
set RS = nothing
set Conn = nothing
Response.Redirect ("AdminEntry.asp?mid=" & menuid & "")
%> 

