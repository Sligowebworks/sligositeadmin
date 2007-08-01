<!--#INCLUDE FILE="adovbs.inc"-->

<%
Set Conn = Server.CreateObject("ADODB.Connection")
	' get conn string from session
	' session("connstring") = "Provider=SQLOLEDB.1;Password=zenmind;Persist Security Info=True;User ID=sa;Initial Catalog=SligoSite_LearnerSupport;Data Source=COLLABITAT"
	' Conn.connectionstring = "Provider=SQLOLEDB.1;Password=zenmind;Persist Security Info=True;User ID=sa;Initial Catalog=SligoSite_LearnerSupport;Data Source=COLLABITAT"
	Conn.connectionstring = Session("connstring")
	Conn.Open 	

 	IF Not IsEmpty(Request.QueryString("PBox")) THEN 
		menuid = request.querystring("PBox")	
	else
		menuid = 0
	end if
	
	sqlcat = "select * from catsort where MainMenuid = " & menuid & " order by sidemenusort;"
		%><%=sqlcat%><br><%
	set rscat = Server.CreateObject("ADODB.Recordset")
	rscat.Open sqlcat, Conn, adOpenKeyset, adLockOptimistic
	
	IF Not IsEmpty(Request.QueryString("CBox")) THEN 
		catid = request.querystring("CBox")	
		FirstCatDefault = "F"
	elseif Not IsEmpty(Request.Form("CategorySelectBox")) then
		%>z<%=Request.Form("CategorySelectBox")%>m<%
		catid = Request.Form("CategorySelectBox")	
		FirstCatDefault = "F"
	else
		%>asdfas<%
		FirstCatDefault = "T"
		catid = rscat("id")
	end if
	
	' 
	sqlmenu = "select * from Menuitems;"
	set rsmenu = Server.CreateObject("ADODB.Recordset")
	rsmenu.Open sqlmenu, Conn, adOpenKeyset, adLockOptimistic
	
			
	SQLset = "select * from Main where Menuid = " & menuid & " and catid = " & catid & " order by CatSort, sort;"
	%><%=sqlset%><%
	set rs = Server.CreateObject("ADODB.Recordset")
	' rs.Open sqlset, Conn, adOpenKeyset, adLockOptimistic
	rs.Open sqlset, Conn, adOpenStatic, adLockreadOnly
	
	SQLsec = "select * from Menuitems where Menuid = " & menuid & " ;"
	set rsm = Server.CreateObject("ADODB.Recordset")
	rsm.Open sqlsec, Conn, adOpenKeyset, adLockOptimistic

	
	
	%>

<script language="javascript">
function ChangeCurrPage()
	{
		zVal = "?PBox=" ;
		zVal = zVal + document.FormPageSelect.PBox.options[document.FormPageSelect.PBox.selectedIndex].value ; 
		document.FormPageSelect.action = "v2AdminMenuMaintenance.asp" + zVal ;
		document.FormPageSelect.submit();
	}
	
function ChangeCategory()
	{
		zVal = "?PBox=" + menuid + "&CBox=" ;
		zVal = zVal + document.FormPageSelect.CBox.options[document.FormPageSelect.CBox.selectedIndex].value ; 
		document.FormPageSelect.action = "v2AdminMenuMaintenance.asp" + zVal ;
		document.FormPageSelect.submit();
	}
	
function ChangeCurrSortUP()
	{
		zVal = "&sid=" ;
		zVal = zVal + document.FormSortSelectUP.PageSelectBoxUP.options[document.FormSortSelectUP.PageSelectBoxUP.selectedIndex].value ; 
		document.FormSortSelectUP.action = "AdminSortProcess.asp?PBox=<%=menuid%>&CBox=<%=catid%>" + zVal ;
		document.FormSortSelectUP.submit();
	}
function ChangeCurrSortDOWN()
	{
		zVal = "&sid=" ;
		zVal = zVal + document.FormSortSelectDOWN.PageSelectBoxDOWN.options[document.FormSortSelectDOWN.PageSelectBoxDOWN.selectedIndex].value ; 
		document.FormSortSelectDOWN.action = "AdminSortChange.asp" + zVal ;
		document.FormSortSelectDOWN.submit();
	}

</script>


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
		<TD><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE"><a HREF="AdminEntry.asp">[Back to Main Menu]</a>
			&nbsp;&nbsp;<FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE"><a HREF="<%=homesource%>">[Jump to home page]</a>
		</TD>
	</TR>
</TABLE>
<br>
	

<TABLE CELLPADDING=3 CELLSPACING=1 ALIGN="CENTER" BORDER="1" BGCOLOR="">
	<TR>
		<TD BGCOLOR="#FFFFCC" COLSPAN = 1 align=middle><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="RED"><B>Section</B></FONT></td>
		<TD BGCOLOR="#FFFFCC" COLSPAN = 1 align=middle><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="RED"><B>Category</B></FONT></td>
	</TR>

<form method=POST name="FormPageSelect">
	<TR>
		<TD>	
			<SELECT	name='PBox' size='1' title='PageSelectTitle' onChange=ChangeCurrPage(<%=rsmenu("menuid")%>)>
				<%
				DO WHILE NOT RSmenu.EOF%>
				
				<%if trim(RSm("MenuItem")) = trim(RSmenu("menuitem")) then
					' ThisIsIt =rsmenu("menuid")
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
	<!--/FORM-->
<!--form method=POST action="v2AdminMenuMaintenance.asp?" name="CategorySelect"-->
		<TD>	
			<SELECT	name='CBox' size='1' title='CategorySelectTitle' onChange=ChangeCategory(<%=rsmenu("menuid")%>)>
					
			<%DO WHILE NOT RScat.EOF
					if trim(RScat("SideMenuText")) = "" or isnull(RScat("SideMenuText")) then
						temper = "same as Section"
					else
						temper = RScat("SideMenuText")
					END IF					
					
					if trim(RScat("id")) = catid then
						
						if TRIM(FirstCatDefault) = "T" then
							ThisIsIt = "same as Section"
						else
							ThisIsIt =rscat("SideMenuText")
						end if	
						strSelected = "SELECTED"
					else
						strSelected = ""
					end if%>
						<OPTION  <%=strSelected%> NAME=MenuSelect value=<%=rscat("id")%>> <%=temper%></option>
					<%rscat.movenext
				loop%>
			</select>
			<%rscat.movefirst%> 
		</TD>
	</tr>
	</FORM>
</TABLE><BR>
<br>
<TABLE CELLPADDING=3 CELLSPACING=1 ALIGN="CENTER" BORDER="1" BGCOLOR="">


	<TR>
		<TD BGCOLOR="#FFFFCC" COLSPAN = 5 align=left><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="RED"><B>SIDE menu  <%=ThisIsIt%></B></FONT></td>
		</TR>
	
	<TR>
		<TD BGCOLOR="#FFFFCC" align=center><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE"><B>name</B></FONT></TD>
		<TD BGCOLOR="#FFFFCC" align=center><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE">sort</FONT></TD>
		<TD BGCOLOR="#FFFFCC" align=center><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE">subsort</FONT></TD>
		<TD BGCOLOR="#FFFFCC" align=center><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE">category</FONT></TD>
		<TD BGCOLOR="#FFFFCC" align=center><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE">show</FONT></TD>
	</TR>

	<% kount = 0	
	
	DO WHILE NOT rs.EOF
	kount = kount + 1%>
   		<TR>
			<TD><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE"><%=trim(RS("Title"))%>&nbsp;</FONT></TD>
			<TD><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE"><%=trim(RS("Catsort"))%>&nbsp;</FONT></TD>
			<TD><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE"><%=trim(RS("sort"))%>&nbsp;</FONT></TD>
			<TD><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE"><%=trim(RS("category"))%>&nbsp;</FONT></TD>
			<% if trim(rs("noshow")) = "False" then
				showit = "yes"
				else
				showit = "no"
				end if%>
			<TD><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE"><%=showit%>&nbsp;</FONT>
			</TD>
			
			<% '<TD><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="RED"><a HREF="AdminSideDelete.asp?id=<%=rs("id")%> <% '&PBox=<%=menuid%> <% ' "> DELETE</FONT></TD> %>
		</TR>
		<%rs.movenext
	loop%> 
</TABLE>
<% 		
if kount > 1 then %>
<br>
<TABLE CELLPADDING=3 CELLSPACING=1 ALIGN="CENTER" BORDER="1" BGCOLOR="">
<form method=POST action="AdminSortProcess.asp?PBox=<%=menuid%>" name="FormSortSelectUP">
	<TR>
		<TD BGCOLOR="#FFFFCC"  align=middle><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="RED"><B>Move up</B></FONT></td>
		<TD BGCOLOR="#FFFFCC"  align=middle><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="RED"><B>Move down</B></FONT></td>
	</TR>
	<TR>
		<TD>	
			<SELECT	name='PageSelectBoxUP' size='1' title='PageSelectTitle' onChange=ChangeCurrSortUP()>
				<OPTION  NAME=MenuSelect value=<%=rsmenu("menuid")%>>Choose to move up</option>
				
				<%rs.movefirst 
				rs.movenext
				DO WHILE NOT RS.EOF%>
					
				<OPTION  NAME=MenuSelect value=<%=rs("id")%>> <%=TRIM(RS("Title"))%></option>
				<%rs.movenext
				loop%>
			</select>
				<%rs.movefirst%> 
				
		</TD>
		</form>
		<form method=POST action="AdminSortProcess.asp?PBox=<%=menuid%>" name="FormSortSelectDOWN">

		<TD>	
			<SELECT	name='PageSelectBoxDOWN' size='1' title='PageSelectTitle' onChange=ChangeCurrSortDOWN()>
				<% rs.movelast 
				' rs.moveprevious
				 temptitle = TRIM(RS("Title"))
				%>
				
				<OPTION  NAME=MenuSelect >Choose to move down</option>
				
				<%  rs.movefirst
				DO WHILE NOT RS.EOF
					IF temptitle <> TRIM(RS("Title")) then%>
						<OPTION  NAME=MenuSelect value=<%=rs("id")%>> <%=TRIM(RS("Title"))%></option>
					<%end if
				rs.movenext
				loop%>
			</select>
				<%rs.movefirst%> 
				
		</TD>
	</tr>
	</FORM>
</TABLE>
<% end if%>



<%rs.close
rsmenu.close
rsm.close		%>
</p>

</body>
</html>

