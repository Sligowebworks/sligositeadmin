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
<%
	Set Conn = Server.CreateObject("ADODB.Connection")
	' get conn string from session
	Conn.connectionstring = Session("connstring")
	Conn.Open 	
'~~~
	IF Not IsEmpty(Request.QueryString("mid")) THEN 
		menuid = request.querystring("mid")	
	else
		menuid = 0
	end if
	
	IF Not IsEmpty(Request.QueryString("idsort")) THEN 
		sorter = request.querystring("idsort")	
	else
		sorter = "N"
	end if
	
	sqlmenu = "select * from Menuitems;"
	
	set rsmenu = Server.CreateObject("ADODB.Recordset")
	rsmenu.Open sqlmenu, Conn, adOpenKeyset, adLockOptimistic
	
'~~~
	IF sorter = "N" then
		SQLset = "select * from Main where Menuid = " & menuid & " order by CatSort, sort;"
	ELSE
		SQLset = "select * from Main where Menuid = " & menuid & " order by ID;"
	END IF
	
	set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open sqlset, Conn, adOpenKeyset, adLockOptimistic
	
'~~~
	SQLsec = "select * from Menuitems where Menuid = " & menuid & " ;"
	set rsm = Server.CreateObject("ADODB.Recordset")
	rsm.Open sqlsec, Conn, adOpenKeyset, adLockOptimistic
	

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Data Entry Form</title>
</head>

<body TEXT="BLUE" ALINK="BLUE" LINK="BLUE">

<% 	logosource = "/SligoSite/" + trim(session("path")) + "/Images/Logo.gif"
	dotsource =  "/SligoSite/" + trim(session("path")) + "/Images/dot.gif"
	homesource = "/SligoSite/" + trim(session("path")) + "/Index.asp"
	proddomain = Session("production domain")
%>

<table WIDTH="600" BORDER="0" align="center">
	<tr>
		<td BGCOLOR="white" ALIGN="middle"><img src="<%=logosource%>" BORDER="0"></td>
	</tr>
	<tr>
		<td BGCOLOR="#0066ff"><IMG height=2 src="<%=dotsource%>"></td>
	</tr>	
</table> 


</TABLE>
<TABLE ALIGN="CENTER">
	<TR>
		<TD><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE">Administrator Main Menu
			&nbsp;&nbsp;<FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE"><a HREF="<%=proddomain%>">[Jump to home page]</a>
			<a HREF="AdminEntry.asp?mid=<%=menuid%>&idsort=Y">[sort on ID]</a>&nbsp;
			<a HREF="AdminEntry.asp?mid=<%=menuid%>&idsort=N">[sort/subsort]</a>
		</TD>
	</TR>
</TABLE>
<br>
<TABLE CELLPADDING=3 CELLSPACING=1 ALIGN="CENTER" BORDER="1" BGCOLOR="">
	<TR>
		<TD BGCOLOR="#FFFFCC" COLSPAN = 3 align=middle><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="RED"><B>Section</B></FONT></td>
	</TR>
	
<script language="javascript">
function ChangeCurrPage()
	{
		zVal = "?mid=" ;
		zVal = zVal + document.FormPageSelect.PageSelectBox.options[document.FormPageSelect.PageSelectBox.selectedIndex].value ; 
		document.FormPageSelect.action = "AdminEntry.asp" + zVal ;
		document.FormPageSelect.submit();
	}
</script>


<form method=POST action="AdminEntry.asp?" name="FormPageSelect">
	<TR>
		<TD>	
			<SELECT	name='PageSelectBox' size='1' title='PageSelectTitle' onChange=ChangeCurrPage()>
				<%DO WHILE NOT RSmenu.EOF%>
					<%if trim(RSm("MenuItem")) = trim(RSmenu("menuitem")) then
					ThisIsIt =rsmenu("menuid")
					strSelected = "SELECTED"
				else
					strSelected = ""
				end if%>
				<OPTION  <%=strSelected%> NAME=MenuSelect value=<%=rsmenu("menuid")%>> <%=RSmenu("menuitem")%></option>
				<%rsmenu.movenext
				loop%>
			</select>
				<%rsmenu.movefirst%> 
				
		</TD>
		<TD><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE"><a HREF="AdminTopEdit.asp?mid=<%=ThisIsIt%>">Edit</FONT></TD>
	</tr>
	</FORM>
	
</TABLE><BR>

<TABLE CELLPADDING=3 CELLSPACING=1 ALIGN="CENTER" BORDER="1" BGCOLOR="">
	<TR>
		<TD BGCOLOR="#FFFFCC" COLSPAN = 2 align=left><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="RED"><B>SIDE menu</B></FONT></td>
		<TD BGCOLOR="#FFFFCC" COLSPAN = 5 align=right><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE"><a HREF="AdminEdit.asp?id=0&mid=<%=menuid%>&Add=Y">Add SIDE item&nbsp;</FONT></TD>
	</TR>
	
	<TR>
		<TD BGCOLOR="#FFFFCC" align=center><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE"><B>edit</B></FONT></TD>
		<TD BGCOLOR="#FFFFCC" align=center><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE">sort</FONT></TD>
		<TD BGCOLOR="#FFFFCC" align=center><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE">subsort</FONT></TD>
		<TD BGCOLOR="#FFFFCC" align=center><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE">category</FONT></TD>
		<TD BGCOLOR="#FFFFCC" align=center><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE">show</FONT></TD>
		<TD BGCOLOR="#FFFFCC" align=center><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLACK">id</FONT></TD>
		<TD BGCOLOR="#FFFFCC" align=center><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="RED">DELETE</FONT></TD>
	</TR>
	<%DO WHILE NOT rs.EOF%>
   		<TR>
			<TD><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE"><a HREF="AdminEdit.asp?id=<%=rs("id")%>&mid=<%=menuid%>"><%=trim(RS("Title"))%>&nbsp;</FONT></TD>
			<TD><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE"><%=trim(RS("Catsort"))%>&nbsp;</FONT></TD>
			<TD><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE"><%=trim(RS("sort"))%>&nbsp;</FONT></TD>
			<TD><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE"><%=trim(RS("category"))%>&nbsp;</FONT></TD>
			<% if trim(rs("noshow")) = "False" then
				showit = "yes"
				else
				showit = "no"
				end if%>
			<TD><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE"><%=showit%>&nbsp;</FONT></TD>
			<TD><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLACK"><%=trim(RS("ID"))%>&nbsp;</FONT></TD>
			<TD><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="RED"><a HREF="AdminSideDeleteQuestion.asp?id=<%=rs("id")%>&mid=<%=menuid%>"> DELETE</FONT></TD>
		</TR>
		<%rs.movenext
	loop%> 
</TABLE>

<table align=center cellpadding=6>
	<TR>
		<TD>
			<% ' <a href="AdminMenuMaintenance.asp?mid=%=ThisIsIt%">Change Menu</a> %>
		</td>
	</tr>
</table>

		
</p>

</body>
</html>

