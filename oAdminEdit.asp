<%
	IF Not IsEmpty(Request.QueryString("mid")) THEN 
		menuid = request.querystring("mid")	
	else
		menuid = 0
	end if
	IF Not IsEmpty(Request.QueryString("id")) THEN 
		id = request.querystring("id")	
	else
		id = 0
	end if
	IF Not IsEmpty(Request.QueryString("Add")) THEN 
		Addit = request.querystring("Add")	
	else
		Addit = "N"
	end if
	
	IF Addit = "D" THEN
 		Response.Redirect ("AdminProcess.asp?mid=" & menuid & "&id=" & id & "&add=" & Addit & "")
	END IF
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.connectionstring = Session("connstring")
	Conn.Open 	
	SQLCount = "select * from Main where id = " & id & " ;"
	' response.write(SQlCount)
	set rs = Server.CreateObject("ADODB.Recordset")
	set rs = conn.execute(SQLCount)
	
	' Conn.connectionstring = "Provider=SQLOLEDB.1;Password=zenmind;Persist Security Info=True;User ID=sa;Initial Catalog=SligoSite_LearnerSupport;Data Source=PARADISO"
	' Conn.Open "learnersupport"
	SQLMenu = "select * from MenuItems where menuid = " & menuid & " ;"
	set rsMenu = conn.execute(SQLMenu)
	
	
	IF Addit = "Y" then	
		title		= ""
		tophead		= ""
		bottomhead 	= ""
		description = ""
		sort		= ""
		category	= ""
		catsort		= ""
		graphic		= ""	
		noshow		= 0
	else
		title		= rs("title")
		tophead		= rs("tophead")
		bottomhead 	= rs("bottomhead")
		description = rs("description")
		sort		= rs("sort")
		category	= rs("category")
		catsort		= rs("catsort")
		graphic		= rs("graphic")	
		noshow		= rs("noshow")
	end if
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

<% 	logosource = "http://Collabitat.com/SligoSite/" + trim(session("path")) + "/Images/Logo.gif"
	dotsource =  "http://Collabitat.com/SligoSite/" + trim(session("path")) + "/Images/dot.gif"
	homesource = "http://Collabitat.com/SligoSite/" + trim(session("path")) + "/Index.asp"
%>
</STYLE>
	<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Data Entry Form</title>
</head>

<body TEXT="BLUE">
<FORM action="AdminProcess.asp" method="post" name="changeform">	
<table WIDTH="720" BORDER="0" align="center">
	<tr>
		<td BGCOLOR="white" ALIGN="middle"><img src="<%=logosource%>" BORDER="0"></td>
	</tr>
	<tr>
		<td BGCOLOR="#0066ff"><IMG height=2 src="<%=dotsource%>"></td>
	</tr>	
</TABLE>
<table WIDTH="720" BORDER="0" align="center">
	<tr>
		<TD><a href="AdminEntry.asp?mid=<%=menuid%>"><FONT FACE=" Arial, Helvetica, sans-serif" SIZE="2">[CANCEL]</FONT></A>
		<TD align="middle"> 
			<FONT FACE=" Arial, Helvetica, sans-serif" SIZE="2">Section/Menu:</FONT>
			&nbsp;<FONT FACE="Arial, Helvetica, sans-serif" color="RED" SIZE="2"><B><%=rsMenu("MenuItem")%>/<%=title%></B></FONT>
		&nbsp;&nbsp;&nbsp;
		<%if Addit = "Y" then%>
			<FONT FACE=" Arial, Helvetica, sans-serif" color="BLUE" SIZE="2"><B>ADDING</B></FONT>
		<%else%>
			<FONT FACE=" Arial, Helvetica, sans-serif" color="BLUE" SIZE="2"><B>EDITING</B></FONT>
		<%end if%>
		
		</TD> 
		<TD colspan=4 ALIGN="CENTER" VALIGN="MIDDLE">	
			<INPUT TYPE=submit Value = "SAVE">&nbsp;&nbsp;<INPUT TYPE=Reset Value = "Restore">
		</TD>
	</tr>	

</table> 
<TABLE WIDTH="725" BORDER="0" BGCOLOR="WHITE" CELLSPACING="3" CELLPADDING="3" ALIGN="CENTER"> 
	<TR>
		<TD HEIGHT="25" align="left" COLSPAN=3>
			<FONT SIZE="2" FACE="Arial, Helvetica, sans-serif">
			<TEXTAREA NAME="description" align="left" WRAP="PHYSICAL" ROWS="20" COLS="70"><%=description%></TEXTAREA> 
		</FONT></TD> 
	</TR>
	</table> 
	<TABLE WIDTH="725" BORDER="0" BGCOLOR="WHITE" CELLSPACING="3" CELLPADDING="3" ALIGN="CENTER"> 
	<TR> 
		<TD>
			<FONT FACE=" Arial, Helvetica, sans-serif" COLOR="BLUE" SIZE="1"><B>Upper title (blank for none)</B></FONT></TD> 
		</TD>
		<TD COLSPAN=2>
			<FONT FACE=" Arial, Helvetica, sans-serif" COLOR="BLUE" SIZE="1"><B>Lower title (blank for none)</B></FONT></TD> 
		</TD>
	</TR>
	<TR>
		<TD WIDTH="334" HEIGHT="25" ><FONT SIZE="2"	FACE="Arial, Helvetica, sans-serif"> 
			<INPUT NAME=tophead SIZE=45 VALUE="<%=tophead%>">
			</FONT></TD> 
		<TD WIDTH="334" HEIGHT="25" COLSPAN=2><FONT SIZE="2"	FACE="Arial, Helvetica, sans-serif"> 
			<INPUT NAME="bottomhead" SIZE="45" VALUE="<%=bottomhead%>">
		</FONT></TD> 
	</TR>
	<TR>
		<TD COLSPAN=3>
			<FONT FACE=" Arial, Helvetica, sans-serif" SIZE="1" COLOR="BLUE" COLSPAN=2><B>Graphic in upper left (blank for none):</B>&nbsp;&nbsp;<INPUT NAME="graphic" SIZE="40" VALUE="<%=graphic%>"></FONT>
		</TD> 
	</TR>
	<TR>
		<TD COLSPAN=3><hr color="green">
			<FONT FACE=" Arial, Helvetica, sans-serif" COLOR="GREEN" SIZE="1" ><B>Side Menu</B></FONT></TD> 
		</TD>
	</tr>
	<TR>
		<TD>
			<FONT FACE=" Arial, Helvetica, sans-serif" COLOR="BLUE" SIZE="1"><B>Side menu category (blank for none)</B></FONT> 
		</TD>
		<TD COLSPAN=2>
			<FONT FACE=" Arial, Helvetica, sans-serif" COLOR="BLUE" SIZE="1" COLSPAN = 2><B>Side menu title (REQUIRED)</B></FONT>
		</TD>
	</tr>	
	<TR> 
		<TD><FONT SIZE="2"	FACE="Arial, Helvetica, sans-serif"> 
			<INPUT NAME="category" SIZE="45" VALUE="<%=Category%>">
		</FONT></TD> 
		<TD COLSPAN = 2><FONT SIZE="2" FACE="Arial, Helvetica, sans-serif"> 
			<INPUT NAME="title" SIZE="45" VALUE="<%=title%>">
		</FONT></TD> 
	</TR>
	<TR> 
		<TD VALIGN="top">
			<FONT FACE=" Arial, Helvetica, sans-serif" SIZE="1" COLOR = "BLUE"><B>Check to hide:</B></FONT>
			&nbsp;&nbsp;<FONT SIZE="2" FACE="Arial, Helvetica, sans-serif">
				<%IF noshow = "True" then%>
					<INPUT TYPE="checkbox" NAME="noshow" CHECKED>
				<%else%>
					<INPUT TYPE="checkbox" NAME="noshow" >
				<%end if%>
			</FONT>
		</TD> 
		<TD>
			<FONT FACE=" Arial, Helvetica, sans-serif" SIZE="1" COLOR="BLUE"><B>Sort:</B></FONT>
			&nbsp;&nbsp;<FONT SIZE="2"	FACE="Arial, Helvetica, sans-serif"> 
			<INPUT NAME="catsort" maxlength=1 SIZE="1" VALUE="<%=catsort%>"> (1-9,a-z)</FONT>
		</TD> 
		<TD>
			<FONT FACE=" Arial, Helvetica, sans-serif" SIZE="1" COLOR="BLUE"><B>Subsort:</B></FONT>
			&nbsp;&nbsp;<FONT SIZE="2"	FACE="Arial, Helvetica, sans-serif"> 
			<INPUT NAME="sort" SIZE="2" maxlength=2 width=2 VALUE="<%=sort%>"> (1-99)</B></FONT>
			</TD> 
	</TR> 
			<!--Dan: These are just the copied over from the safeops page; there's no coding yet.-->
			<input type=hidden value="<%=id%>" name="id">
			<input type=hidden value="<%=menuid%>" name="mid">
			<input type=hidden value=<%=Addit%> name="Addit">
	</TABLE>
	
	</FORM>

</body>
</html>

