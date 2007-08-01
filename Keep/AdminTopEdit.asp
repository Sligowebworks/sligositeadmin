<!--#INCLUDE FILE="adovbs.inc"-->
<%
	IF Not IsEmpty(Request.QueryString("mid")) THEN 
		menuid = request.querystring("mid")	
	else
		menuid = 0
	end if
	
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.connectionstring = Session("connstring")
	Conn.Open 	
	set rs = Server.CreateObject("ADODB.Recordset")
	
	'	Conn.Open "learnersupport"
	sql = "select * from MenuItems where menuid = " & menuid & " ;"
	rs.Open sql, Conn, adOpenKeyset, adLockOptimistic
	
	' set rs = conn.execute(SQL)
	menuitem	= rs("menuitem")
	
	%>
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
	<title>Data Entry Form</title>
</head>

<body TEXT="BLUE">

<table WIDTH="720" BORDER="0" align="center">
	<tr>
		<td BGCOLOR="white" ALIGN="middle"><img src="<%=logosource%>" BORDER="0"></td>
	</tr>
	<tr>
		<td BGCOLOR="#0066ff"><IMG height=2 src="Images/Dot.gif"></td>
	</tr>	
</table>
<TABLE ALIGN="CENTER">
	<TR><TD>
		<FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE">
		&nbsp;&nbsp;Administrator pages: Main menu > Section EDIT</font>	
	</TD></TR>
</TABLE>
 
<TABLE ALIGN="CENTER">
<FORM action="AdminTopProcess.asp" method="post" name="changeform">	
	<tr>
		<TD align="middle"> 
			<FONT FACE="Arial, Helvetica, sans-serif" color="RED" SIZE="2">
			<a HREF="AdminEntry.asp?mid=<%=MenuID%>">[Cancel]</a>
			&nbsp;&nbsp;<b><%=menuItem%></b></FONT>
			&nbsp;&nbsp;&nbsp;

		<input type=hidden value="<%=menuid%>" name="mid">
		<input type=hidden value="<%=session("connstring")%>" name="conns">
			&nbsp;&nbsp;<INPUT TYPE=submit Value = "Save">
			&nbsp;&nbsp;<INPUT TYPE=Reset Value = "Restore">
			
			</TD>
		</TD> 
	</tr>	

</table> 
<br>
<TABLE WIDTH="600" BORDER="0" BGCOLOR="WHITE" CELLSPACING="3" CELLPADDING="3" ALIGN="CENTER"> 
	<TR> 
		<TD>
			<FONT FACE=" Arial, Helvetica, sans-serif" SIZE="2"><B>Section:</B></FONT>
		</TD> 
		<TD WIDTH="334" HEIGHT="25" colspan=3><FONT SIZE="2" FACE="Arial, Helvetica, sans-serif"> 
			<INPUT NAME="menuitem" SIZE="45" VALUE="<%=menuitem%>">
			</FONT>
		</TD> 
	</TR> 
	<TR> 
		<TD>
			<FONT FACE=" Arial, Helvetica, sans-serif" SIZE="2"><B>Left/right sort:</B></FONT>
		</TD> 
		<TD WIDTH="334" HEIGHT="25" colspan=3><FONT SIZE="2" FACE="Arial, Helvetica, sans-serif"> 
			<%=menuid%>
			<INPUT type=HIDDEN NAME="menuid" VALUE="<%=menuid%>" >
			</FONT>
		</TD> 
	</TR> 
	</TABLE> 
	</FORM>

</body>
</html>

