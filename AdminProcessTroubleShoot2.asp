<!--#INCLUDE FILE="adovbs.inc"-->
<STYLE type="text/css">
<!--
H3{color:WHITE;font-size:8pt;font-family:Verdana, Arial, Helvetica, sans-serif;}
TD, #basicstyle{color:black;font-size:9pt;font-family:Verdana, Arial, Helvetica, sans-serif;text-decoration: none;}
#tophead{font-family:Verdana, Arial, Helvetica; font-size:9pt; color:#990033;font-weight:BOLD;}
#bottomhead{font-family:times new roman; font-size:16pt; color:#999999; font-style:italic;}
#cathead{font-family:Verdana, Arial, Helvetica; font-size:11pt; color:#FFFFCC;font-weight:none;}
#colhead{font-family:Verdana, Arial, Helvetica; font-size:11pt; color:#FFFFCC;font-weight:none;}
#sidelink{font-family:Verdana, Arial, Helvetica; font-size:8pt; color:#FFFFCC;font-weight:none;}
#redlink{font-family:Verdana, Arial, Helvetica; font-size:8pt; color:#990033;font-weight:bold;}
#topredlink{color:#990033;font-size:8pt;font-family:Verdana, Arial, Helvetica, sans-serif;font-weight:BOLD; text-decoration: none;} 
A:Link{color:NAVY;font-size:8pt;font-family:Verdana, Arial, Helvetica, sans-serif;font-weight:BOLD; text-decoration: none;} 
A:Visited{color:NAVY;font-size:8pt;font-family:Verdana, Arial, Helvetica, sans-serif;font-weight:BOLD; text-decoration: none;}  
-->

<% 	logosource = "/SligoSite/" + trim(session("path")) + "/Images/Logo.gif"
	dotsource =  "/SligoSite/" + trim(session("path")) + "/Images/dot.gif"
	homesource = "/SligoSite/" + trim(session("path")) + "/Index.asp"
%>
</STYLE>
<html>
<head>
	<title>Data Entry Form</title>
</head>

<body TEXT="BLUE">
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<table WIDTH="720" BORDER="0" align="center">
	<tr>
		<td BGCOLOR="white" ALIGN="middle"><img src="<%=logosource%>" BORDER="0"></td>
	</tr>
	<tr>
		<td BGCOLOR="#0066ff"><IMG height=2 src="<%=dotsource%>"></td>
	</tr>	
</TABLE>
 <%
    ' IF Not IsEmpty(Request.form("id")) THEN 
	id = Request.Form("id")	
	menuid = Request.Form("mid")
	response.Write(menuid)
	'else
	'	id = 0
	'end if
       	if Request.Form("title") = "" or IsEmpty(Request.Form("title")) then
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
'create connection objects
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.connectionstring = Session("connstring")
Conn.Open 	
'-------------------------------------------------------------------------------------------------------------------
'verify page name and redirect user back to Admin Entry page
if (Addit="Y") then
    sqlverify="SELECT COUNT(*) AS counter FROM Main WHERE pagename='"&PageName &"' AND menuname='"&menuName&"';"
    'response.Write sqlverify
    set rsverify=server.CreateObject("ADODB.Recordset")
    rsverify.Open sqlverify,Conn,adOpenKeyset,adLockOptimistic
    while not rsverify.EOF 
        counter=rsverify.Fields("counter")
        rsverify.MoveNext
    wend
    if counter>0 then 
        response.Redirect("verifypage.asp")
    end if
    set rsverify=nothing
end if
'-------------------------------------------------------------------------------------------------------------------
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
'	vfile = "D:\Websites\SligoSite\CDHP\" & Session("path") & "\" & RS("menuname") & "\"  & RS("PageName") & ".asp"
    vfile = "D:\Websites\SligoSite\CDHP\" & RS("menuname") & "\"  & RS("PageName") & ".asp"
 '   response.Write(vfile)
'	response.write("d:\websites\sligosite\" & Session("path") & "\" & "template.asp" & "<BR>")
'	response.write(vfile)
	objFSO.CopyFile "D:\Websites\SligoSite\CDHP\Template\" &  "template.asp",vfile
end if
Conn.Close
' free memory
set RS = nothing
set Conn = nothing
Response.Redirect ("AdminEntry.asp?mid=" & menuid & "")
%> 

</body>
</html>