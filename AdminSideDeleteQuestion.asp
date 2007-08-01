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
</STYLE>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>Data Entry Delete Form</title>
</head>
<body TEXT="BLUE" ALINK="BLUE" LINK="BLUE">
<% 	logosource = "/SligoSite/" + trim(session("path")) + "/Images/Logo.gif"
	dotsource =  "/SligoSite/" + trim(session("path")) + "/Images/dot.gif"
	homesource = "/SligoSite/" + trim(session("path")) + "/Index.asp"
%>

<table WIDTH="600" BORDER="0" align="center">
	<tr>
		<td BGCOLOR="white" ALIGN="middle"><img src="<%=logosource%>" BORDER="0"></td>
	</tr>
	<tr>
		<td BGCOLOR="#0066ff"><IMG height=2 src="<%=dotsource%>"></td>
	</tr>	
</TABLE>
<TABLE ALIGN="CENTER">
	<TR>
		<TD><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE">Administrator Delete Form
			&nbsp;&nbsp;<FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE"><a HREF="AdminEntry.asp">[Cancel and return to Admin page]</a>
		</TD>
	</TR>
</TABLE>
 <%
	IF Not IsEmpty(Request.QueryString("mid")) THEN 
		menuid = request.querystring("mid")	
	else
		menuid = 0
	end if
	
IF Not IsEmpty(Request.QueryString("id")) THEN 
	id = request.querystring("id")	
	
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.connectionstring = Session("connstring")
	Conn.Open 	
	SQL = "select * from Main where id = " & id & " ;"
	set rs = Server.CreateObject("ADODB.Recordset")
	set rs = conn.execute(SQL)
	

		%>
	<table WIDTH="600" BORDER="0" align="center">
		<tr>
			<TD align = center>
			<br>Really delete <%=rs("title")%>?<br><br>
			
			<form method=POST action="AdminDelete.asp?id=<%=id%>" name="FormPageSelect">
			<INPUT NAME=ConfirmDelete value="Confirm Delete" type="Submit">
			</FORM>
			</TD>
		</TR>
	</TABLE>
	</body>
	</html>
		<%
	else
		id = 0
		Response.Redirect ("AdminEntry.asp?mid=" & menuid & "")
	end if%>