<!--#INCLUDE FILE="adovbs.inc"-->
<%
Set Conn = Server.CreateObject("ADODB.Connection")
' get conn string from session
Conn.connectionstring = Session("connstring")
Conn.Open 	
BoxUP=Request.Form("BoxUP")
BoxDOWN=Request.Form("BoxDOWN")

IF Not IsEmpty(Request.Form("PBox")) THEN 
	menuid = request.Form("PBox")	
	' see if there's something to move up or down.
	if (not IsEmpty(Request.Form("BoxUP")) and Request.Form("BoxUP") <> 0) or (not IsEmpty(Request.Form("BoxDOWN")) and Request.Form("BoxDOWN") <> "0") then
		
		'get id, catid and sort from main
		sqlmover = "select * from main where Menuid = " & menuid & " and catid = " & Request.Form("CBox") & " order by [sort];"
		set rsmover = Server.CreateObject("ADODB.Recordset")
		rsmover.Open sqlmover, Conn, adOpenStatic, adLockOptimistic
		' store sort for the changer and changee
		IF not IsEmpty(Request.Form("BoxUP")) and Trim(Request.Form("BoxUP")) <> "0" then	
			'movenext and grab sort for changee
			' rsmover.find("sort=98")
			do while not rsmover.eof
				abovesort = rsmover("sort")
				aboveid = rsmover("id")
				rsmover.movenext
				if trim(rsmover("id")) = trim(BoxUP) then
					exit do
				end if
			loop
			thissort = rsmover("sort")
			rsmover.close
			' update sort for one above  
			sqlchangetomove = "update main set sort = " & abovesort & " where id = " & BoxUP & ";"	
			set rsmove1 = Server.CreateObject("ADODB.Recordset")
			rsmove1.Open sqlchangetomove, Conn, adOpenKeyset, adLockPessimistic
			'and one picked
			sqlchangetomove = "update main set sort = " & thissort & " where id = " & aboveid & ";"	
			set rsmove2 = Server.CreateObject("ADODB.Recordset")
			rsmove2.Open sqlchangetomove, Conn, adOpenKeyset, adLockPessimistic
			' rsmove2.close
			' rsmove1.close	
					
		elseIF not IsEmpty(Request.Form("BoxDOWN")) and Trim(Request.Form("BoxDOWN")) <> "0" then	
			'MOVE until you go past and grab sort for changee
			do while not rsmover.eof
				thissort = rsmover("sort")
				if trim(rsmover("id")) = trim(BoxDOWN) then
					rsmover.movenext
					exit do
				end if
				rsmover.movenext
			loop
			belowsort = rsmover("sort")
			belowid = rsmover("id")
			rsmover.close
			' update sort for one below  
			sqlchangetomove = "update main set sort = " & belowsort & " where id = " & BoxDOWN & ";"	
			set rsmove1 = Server.CreateObject("ADODB.Recordset")
			rsmove1.Open sqlchangetomove, Conn, adOpenKeyset, adLockPessimistic
			'and one picked
			sqlchangetomove = "update main set sort = " & thissort & " where id = " & belowid & ";"	
			set rsmove2 = Server.CreateObject("ADODB.Recordset")
			rsmove2.Open sqlchangetomove, Conn, adOpenKeyset, adLockPessimistic

		end if ' up or down
	end if 'sid
else
	menuid = 0
end if ' PBox

	
sqlcat = "select * from catsort where MainMenuid = " & menuid & " order by sidemenusort;"
set rscat = Server.CreateObject("ADODB.Recordset")
rscat.Open sqlcat, Conn, adOpenKeyset, adLockOptimistic

IF Not IsEmpty(Request.Form("CBox")) THEN 

'check to see that catid exists for this menuid
sqlcheckcatid = "select * from catsort where MainMenuid = " & menuid & " order by sidemenusort;"
set rscatcheck = Server.CreateObject("ADODB.Recordset")
rscatcheck.Open sqlcat, Conn, adOpenKeyset, adLockOptimistic
gotem = "F"
do while not rscatcheck.eof
	if trim(Request.Form("CBox")) = trim(rscatcheck("id")) then
		gotem = "T"
	end if
rscatcheck.movenext
loop
if gotem = "T" then
	catid = request.Form("CBox")	
	' FirstCatDefault = "F"
else
	' 	FirstCatDefault = "T"
	catid = rscat("id")
end if
rscatcheck.close
else
' FirstCatDefault = "T"
catid = rscat("id")
end if

' 
sqlmenu = "select * from Menuitems;"
set rsmenu = Server.CreateObject("ADODB.Recordset")
rsmenu.Open sqlmenu, Conn, adOpenKeyset, adLockOptimistic

	
SQLset = "select * from Main where Menuid = " & menuid & " and catid = " & trim(catid) & " order by sort;"
set rs = Server.CreateObject("ADODB.Recordset")
' rs.Open sqlset, Conn, adOpenKeyset, adLockOptimistic
rs.Open sqlset, Conn, adOpenStatic, adLockreadOnly

kount = rs.recordcount

SQLsec = "select * from Menuitems where Menuid = " & menuid & " ;"
set rsm = Server.CreateObject("ADODB.Recordset")
rsm.Open sqlsec, Conn, adOpenKeyset, adLockOptimistic
%>

<script language="javascript">
	
function ChangeCurrSortUP()
	{
		zVal = "&sid=" ;
		zVal = zVal + document.FormSortSelectUP.PageSelectBoxUP.options[document.FormSortSelectUP.PageSelectBoxUP.selectedIndex].value ; 
		document.FormSortSelectUP.action = "v2AdminMenuMaintenance.asp?PBox=<%=menuid%>&CBox=<%=catid%>&Dir=U" + zVal ;
		document.FormSortSelectUP.submit();
	}
function ChangeCurrSortDOWN()
	{
		zVal = "&sid=" ;
		zVal = zVal + document.FormSortSelectDOWN.PageSelectBoxDOWN.options[document.FormSortSelectDOWN.PageSelectBoxDOWN.selectedIndex].value ; 
		document.FormSortSelectDOWN.action = "v2AdminMenuMaintenance.asp?PBox=<%=menuid%>&CBox=<%=catid%>&Dir=D" + zVal ;
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
		<TD><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE"><a HREF="v2AdminEntry.asp">[Back to Main Menu]</a>
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

<form method=POST name="FormSelect">
	<TR>
		<TD>	
			<SELECT	name='PBox' size='1' title='PBoxTitle' onChange="document.FormSelect.action='v2AdminMenuMaintenance.asp';document.FormSelect.submit();">
				<%
				DO WHILE NOT RSmenu.EOF%>
				
				<%if trim(RSm("MenuItem")) = trim(RSmenu("menuitem")) then
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
		<TD>	
			<SELECT	name='CBox' size='1' title='CBoxTitle' onChange="document.FormSelect.action='v2AdminMenuMaintenance.asp';document.FormSelect.submit();">
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
						<OPTION  <%=strSelected%> NAME=CategorySelect value=<%=rscat("id")%>> <%=temper%></option>
				<%
				rscat.movenext
				loop%>
			</select>
			<%rscat.movefirst%> 
		</TD>
	</tr>
</TABLE>
<br>
<%if kount > 1 then %>
<TABLE CELLPADDING=3 CELLSPACING=1 ALIGN="CENTER" BORDER="1" BGCOLOR="">
	<TR>
		<TD BGCOLOR="#FFFFCC"  align=middle><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="RED"><B>Move up</B></FONT></td>
		<TD BGCOLOR="#FFFFCC"  align=middle><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="RED"><B>Move down</B></FONT></td>
	</TR>
	<TR>
		<TD>	
			<SELECT	name='BoxUP' size='1' title='PageSelectTitleUP' onChange="document.FormSelect.action='v2AdminMenuMaintenance.asp';document.FormSelect.submit();">
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
		<TD>	
			<SELECT	name='BoxDOWN' size='1' title='PageSelectTitleDOWN' onChange="document.FormSelect.action='v2AdminMenuMaintenance.asp';document.FormSelect.submit();">
				<% rs.movelast 
				temptitle = TRIM(RS("Title"))%>
				<OPTION  NAME=MenuSelect value=<%=rsmenu("menuid")%>>Choose to move down</option>
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
</TABLE>
	</FORM>
<%Else%>
	</FORM>
<%end if%>

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

	<% ' kount = 0	
	
	DO WHILE NOT rs.EOF%>
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
<%rs.close
rsmenu.close
rsm.close%>
</p>

</body>
</html>
<form method=POST action="AdminSortProcess.asp?PBox=<%=menuid%>" name="FormSortSelectUP">
</form>
<form method=POST action="AdminSortProcess.asp?PBox=<%=menuid%>" name="FormSortSelectDOWN">
		
