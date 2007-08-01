	<% 
If Request.QueryString("SecID") <> "" then
	SecID = Request.QueryString("SecID")
Else
	SecID = 0
End IF	
	
If Request.QueryString("Type") <> "" then
	SSType = Request.QueryString("Type")
Else
	SSType = "I"
End IF

If Request.QueryString("CatID") <> "" then
	CatID = Request.QueryString("CatID")
Else
	CatID = 0
End IF

If Request.QueryString("ID") <> "" then
	SSID = Request.QueryString("ID")
Else
	SSID = "I"
End IF

'Select Case Ucase(SSType)
'	case "I"
'	case "C"
'	case "S"
'End Select

	 %>
	<% ' Changed Catsort 
	  IF Not IsEmpty(Request.QueryString("CBox")) THEN 
		CBox = request.querystring("CBox")	
	else
		CBox = 0
	end if
	if CBox <> 0 then
		sqlcatchange = "update Main set catid= (" & request.form("CatsortSelectBox") & ") where id=" & CBox & ";"
		set rscatchange = Server.CreateObject("ADODB.Recordset")

		rscatchange.Open sqlcatchange, Conn, adOpenKeyset, adLockPessimistic
	end if
	%>
	
<%  IF Not IsEmpty(Request.QueryString("SecID")) THEN 
		SecID = request.querystring("SecID")	
	else
		SecID = 0
	end if
	
	sqlmenu = "select * from Menuitems;"
	set rsmenu = Server.CreateObject("ADODB.Recordset")
	rsmenu.Open sqlmenu, Conn, adOpenKeyset, adLockOptimistic
	
	SQLset = "select * from Main where Menuid = " & SecID & " order by Catid, sort;"
	set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open sqlset, Conn, adOpenKeyset, adLockOptimistic
		
	'SELECTED SECTION
	'UNNECESSARY????  USE QUERYSTRING INSTEAD?
	SQLsec = "select * from Menuitems where Menuid = " & SecID & " ;"
	set rsm = Server.CreateObject("ADODB.Recordset")
	rsm.Open sqlsec, Conn, adOpenKeyset, adLockOptimistic
	
	'mzd SQL NOT FINISHED -- SELECTED CATEGORY
	'SQLcat = "select * from Menuitems where Menuid = " & SecID & " ;"
	SQLcat = "select * from catsort where MainMenuid = " & SecID & " order by sidemenusort;"
	set rsc = Server.CreateObject("ADODB.Recordset")
	rsc.Open SQLcat, Conn, adOpenKeyset, adLockOptimistic
	Response.Write SQLcat
	%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Data Entry Form</title>
</head>

<body TEXT="BLUE" ALINK="BLUE" LINK="BLUE">

<% 	logosource = "http://Collabitat.com/SligoSite/" + trim(session("path")) + "/Images/Logo.gif"
	dotsource =  "http://Collabitat.com/SligoSite/" + trim(session("path")) + "/Images/dot.gif"
	homesource = "http://Collabitat.com/SligoSite/" + trim(session("path")) + "/Index.asp"
%>

<!-- <table WIDTH="600" BORDER="0" align="center">
	<tr>
		<td BGCOLOR="white" ALIGN="middle"><img src="<%=logosource%>" BORDER="0"></td>
	</tr>
	<tr>
		<td BGCOLOR="#0066ff"><IMG height=2 src="<%=dotsource%>"></td>
	</tr>	
</table>  -->


</TABLE>
<TABLE ALIGN="CENTER">
	<TR>
		<TD><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE">Administrator Main Menu
			&nbsp;&nbsp;<FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE"><a HREF="<%=homesource%>">[Jump to home page]</a>
		</TD>
	</TR>
</TABLE>
<br>
<script language="javascript">
function ChangeCurrPage(zSelect)
	{	
		zVal = "?SecID=" ;
		zVal = zVal + zSelect.options[zSelect.selectedIndex].value ; 
		document.Form.action = '<%= zThisPage %>' + zVal ;
		document.Form.submit();
		
		/* old code
		zVal = "?SecID=" ;
		zVal = zVal + document.FormPageSelect.PageSelectBox.options[document.FormPageSelect.PageSelectBox.selectedIndex].value ; 
		document.FormPageSelect.action = "v2AdminEntry.asp" + zVal ;
		document.FormPageSelect.submit();
		*/
	}
	
</script>
<form method="post" action="<%= zThisPage %>" name="Form">
<TABLE CELLPADDING=3 CELLSPACING=1 ALIGN="CENTER" BORDER="1" BGCOLOR="">
	<TR>
		<TD BGCOLOR="#FFFFCC" COLSPAN="3" align="center"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="RED"><B>Sections</B></FONT></td>
	</TR>
	<!-- <form method=POST action="v2AdminEntry.asp?" name="FormPageSelect"> -->
	<TR>
		<TD>	
		<!-- sections -->
			<SELECT	name="Sections" size="1" onChange="ChangeCurrPage(document.Form.Sections)">
				<%DO WHILE NOT RSmenu.EOF%>
					<%if trim(RSm("MenuItem")) = trim(RSmenu("menuitem")) then
					ThisIsIt = rsmenu("menuid")
					strSelected = "SELECTED"
				else
					strSelected = ""
				end if%>
				<OPTION  <%=strSelected%> value="<%=rsmenu("menuid")%>"> <%=RSmenu("menuitem")%></option>
				<%rsmenu.movenext
				loop%>
			</select>
				<%rsmenu.movefirst%> 
				
		</TD>
		<TD><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE"><a HREF="v2AdminTopEdit.asp?SecID=<%=ThisIsIt%>">Edit</FONT></TD>
	</tr>
	<!-- </FORM> -->
	
</TABLE><BR>
<TABLE CELLPADDING=3 CELLSPACING=1 ALIGN="CENTER" BORDER="1" BGCOLOR="">
	<TR>
		<TD BGCOLOR="#FFFFCC" COLSPAN = 3 align=middle><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="RED"><B>Categories</B></FONT></td>
	</TR>
	<!-- <form method=POST action="v2AdminEntry.asp?" name="FormPageSelect"> -->
	<TR>
		<TD>	
		<!-- categories -->
			<SELECT	name="categories" size="1" onChange="ChangeCurrPage('document.Form.Categories')">
				<%DO WHILE NOT rsc.EOF%>
					<%if CatID = rsc("id") then
					ThisIsIt = rsc("id")
					strSelected = "SELECTED"
				else
					strSelected = ""
				end if%>
				<OPTION  <%=strSelected%> value="<%=rsc("ID")%>"> <%=rsc("SideMenuText")%></option>
				<%rsc.movenext
				loop%>
			</select>
				<%rsc.movefirst%> 
				
		</TD>
		<TD><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE"><a HREF="v2AdminTopEdit.asp?SecID=<%=ThisIsIt%>">Edit</FONT></TD>
	</tr>
	<!-- </FORM> -->
	
</TABLE><BR>


<TABLE CELLPADDING=3 CELLSPACING=1 ALIGN="CENTER" BORDER="1" BGCOLOR="">
	<TR>
		<TD BGCOLOR="#FFFFCC" COLSPAN = 2 align=left><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="RED"><B>SIDE items</B></FONT></td>
		<TD BGCOLOR="#FFFFCC" COLSPAN = 3 align=right><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE"><a HREF="v2AdminEdit.asp?id=0&SecID=<%=SecID%>&Add=Y">Add SIDE item&nbsp;</FONT></TD>
	</TR>
	
	<TR>
		<TD BGCOLOR="#FFFFCC" align=center><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE"><B>edit</B></FONT></TD>
		<TD BGCOLOR="#FFFFCC" align=center><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE">category</FONT></TD>
		<TD BGCOLOR="#FFFFCC" align=center><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE">subsort</FONT></TD>
		<TD BGCOLOR="#FFFFCC" align=center><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE">category</FONT></TD>
		<TD BGCOLOR="#FFFFCC" align=center><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE">show</FONT></TD>
		</TR>
	<%DO WHILE NOT rs.EOF
	
		SQLcs = "select * from catsort where MainMenuid = " & SecID & " order by SideMenuSort;"
		set rscs = Server.CreateObject("ADODB.Recordset")
		rscs.Open sqlcs, Conn, adOpenKeyset, adLockOptimistic
		%>
		<!-- <form method=POST action="v2AdminEntry.asp?" name="x<%=rs("id")%>"> -->
		<TR>
			<TD><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE"><a HREF="v2AdminEdit.asp?id=<%=rs("id")%>&SecID=<%=SecID%>"><%=trim(RS("Title"))%>&nbsp;</FONT></TD>
			<TD><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE">			
			<SELECT	name='CatsortSelectBox<%=rs("id")%>' size='1' title='CatsortSelectTitle' onChange="document.Form.action='v2AdminEntry.asp?SecID=<%=SecID%>&CBox=<%=rs("id")%>'; document.Form.submit();">
				<%DO WHILE NOT rscs.EOF
					if rs("catid") = rscs("id") then
						ThisIsItcs =rscs("id")
						strSelected = "SELECTED"
					else
						strSelected = ""
					end if%>
	
					<OPTION  <%=strSelected%> value="<%=rscs("id")%>"><%=RScs("SideMenuText")%></option>
					<%rscs.movenext
				loop%>
						
			</select>
					
				<%rscs.movefirst%> 
				
			<% ' =trim(RS("Catsort"))%>
			
			
			&nbsp;</FONT></TD>
			<TD><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE"><%=trim(RS("sort"))%>&nbsp;</FONT></TD>
			<TD><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE"><%=trim(RS("category"))%>&nbsp;</FONT></TD>
			<% if trim(rs("noshow")) = "False" then
				showit = "yes"
				else
				showit = "no"
				end if%>
			<TD><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE"><%=showit%>&nbsp;</FONT></TD>
			<% '<TD><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="RED"><a HREF="v2AdminSideDelete.asp?id=<%=rs("id")%> <% '&SecID=<%=SecID%> <% ' "> DELETE</FONT></TD> %>
		</TR>
		
		<%rs.movenext
	loop%> 
	
</TABLE>
</FORM>
<table align=center cellpadding=6>
	<TR>
		<TD>
			<a href="v2AdminMenuMaintenance.asp?SecID=<%=ThisIsIt%>">Change Menu</a>
		</td>
	</tr>
</table>

		
</p>

</body>
</html>

