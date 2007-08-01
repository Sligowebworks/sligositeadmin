 <!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!-- copyright 2001 Sligo Computer Services, Inc. All rights reserved -->
<html>
<head>
<title>Sligo Computer Services, Inc. Home Page</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<meta name="keywords" content="computer services, web design, software development, programming, consultants, visual basic,
	school improvement, assessment, Maryland, D.C. Metropolitan area, VA, statistical analysis, data analysis, visual data presentation, networks, desktop publishing, layout,sligo">
	<meta name="description" content="Sligo Computer Services Web Site">	
</HEAD>
<script type="text/javascript" language="JavaScript">
function opennew(zurl) {
window.open(zurl, "newwindow", "status,Height=400,WIDTH=750,alwaysRaised,resizable,scrollbars,screenx=100,screeny=50")
}
</script>
<%	Set sp = Server.CreateObject("SligoSite.Functions")
	'response.write sp.qstr
	'mzd -- move to global.asa onSessionStart
	Session("strConn") = "Provider=SQLOLEDB.1;Password=zenmind;Persist Security Info=True;User ID=sa;Initial Catalog=SligoSite_Sligo;Data Source=COLLABITAT"
	sp.InitializeSligoSite
%>
<STYLE type="text/css">
<!--
TD, #basicstyle{color:black;font-size:9pt;font-family:Verdana, Arial, Helvetica, sans-serif;text-decoration: none;}
#bottomhead{font-family:times new roman; font-size:16pt; color:#999999; font-style:italic;}
#tophead{font-family:Verdana, Arial, Helvetica; font-size:9pt; color:#333333;font-weight:BOLD;}
#Name{font-family:times new roman; font-size:16pt; color:#999999; font-style:italic;}
#cathead{font-family:Verdana, Arial, Helvetica; font-size:11pt; color:#333333;font-weight:bold;}
#colhead{font-family:Verdana, Arial, Helvetica; font-size:11pt; color:#333333;font-weight:bold;}
#redlink{font-family:Verdana, Arial, Helvetica; font-size:8pt; color:#990033;font-weight:bold;}
#topredlink{font-family:Verdana, Arial, Helvetica, sans-serif; font-size:8pt; color:#990033;font-weight:bold;}
#sidelink{color:#229999;font-family:Verdana, Arial, Helvetica; font-size:8pt;font-weight:none;}
A:Link{color:#229999;font-size:8pt;font-family:, Verdana, Arial, Helvetica, sans-serif;font-weight:bold; text-decoration: none;} 
A:Visited{color:#229999;font-size:8pt;font-family:Verdana, Arial, Helvetica, sans-serif;font-weight:BOLD; text-decoration: none;}  -->
</STYLE>
<DIV ALIGN="CENTER">
<body BGCOLOR="white" LINK="#229999" ALINK="#229999" VLINK="#229999"><!--BACKGROUND="Images/bg4.gif">-->
<% ' The table for the image and top menu starts here %>
<table WIDTH="600">
	<tr>
		<td ALIGN="middle" COLSPAN="2" BGCOLOR="white">
		<IMG ALIGN="left" SRC="Images/Logo.gif">
		<IMG ALIGN="right" SRC="Images/sligocollage.gif"></td>
	</tr>	
	<tr>
		<%' keep out for now . . . <td ALIGN="LEFT"><font SIZE="1pt" FACE="Verdana, Arial, Helvetica"><A HREF="http://www.swells.com/todolist/" TARGET="NEW" NAME="IPM">IPM</A>&nbsp;&nbsp;<A HREF="http://www.collabitat.com/Rolodex.asp" TARGET="NEW" NAME="Rolodex">rolodex</FONT></A></td>%>
		<td  ALIGN="right" VALIGN="bottom">
		<P id="topredlink">
			<%sp.PaintTopMenu()%>
		</P>		
		</td>		 
	</tr>
	<tr>
		<td COLSPAN="2" BGCOLOR="#cccccc"><IMG height=2 src="Images/Dot.gif"></td>
	</tr>
</table>
<% ' now for the bottom of the page. First the big table %>
<table WIDTH="600" cellpadding="5" BORDER="0">
	<tr>
		<td WIDTH="145" BGCOLOR="#ffffff" HEIGHT="450" ALIGN="left" VALIGN="top" >
		<TABLE WIDTH="140" ALIGN="left">
		<% ' This is the side menu table%>
			<tr>
				<td>
					<P id="colhead">		
						<%sp.PaintSideTitle()%>
						            
						<br>
					</P>
				</td>
			</tr>
						<%sp.PaintSideMenu()%>	
		</TABLE>
		</td>
		<% ' Now comes the area for content %>
		<td width=430 VALIGN="top" BACKGROUND="Images/bg1.gif" >
			<TABLE ALIGN="JUSTIFY" WIDTH="430" CELLPADDING="5">
				<TR><TD>	
<%
'*******************************************
'sp.PaintContent() %>
					
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
		<TD><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE">Administrator Main Menu
			&nbsp;&nbsp;<FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE"><a HREF="<%=homesource%>">[Jump to home page]</a>
		</TD>
	</TR>
</TABLE>
<br>
<script language="javascript">
function ChangeCurrPage()
	{
		zVal = "?PBox=" ;
		zVal = zVal + document.FormPageSelect.PageSelectBox.options[document.FormPageSelect.PageSelectBox.selectedIndex].value ; 
		document.FormPageSelect.action = "v2AdminEntry.asp" + zVal ;
		document.FormPageSelect.submit();
	}
	
</script>

<TABLE CELLPADDING=3 CELLSPACING=1 ALIGN="CENTER" BORDER="1" BGCOLOR="">
	<TR>
		<TD BGCOLOR="#FFFFCC" COLSPAN = 3 align=middle><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="RED"><B>Sections</B></FONT></td>
	</TR>
	<form method=POST action="v2AdminEntry.asp?" name="FormPageSelect">
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
		<TD><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE"><a HREF="v2AdminTopEdit.asp?PBox=<%=ThisIsIt%>">Edit</FONT></TD>
	</tr>
	</FORM>
	
</TABLE><BR>
<TABLE CELLPADDING=3 CELLSPACING=1 ALIGN="CENTER" BORDER="1" BGCOLOR="">
	<TR>
		<TD BGCOLOR="#FFFFCC" COLSPAN = 3 align=middle><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="RED"><B>Categories</B></FONT></td>
	</TR>
	<form method=POST action="v2AdminEntry.asp?" name="FormPageSelect">
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
		<TD><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE"><a HREF="v2AdminTopEdit.asp?PBox=<%=ThisIsIt%>">Edit</FONT></TD>
	</tr>
	</FORM>
	
</TABLE><BR>


<TABLE CELLPADDING=3 CELLSPACING=1 ALIGN="CENTER" BORDER="1" BGCOLOR="">
	<TR>
		<TD BGCOLOR="#FFFFCC" COLSPAN = 2 align=left><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="RED"><B>SIDE items</B></FONT></td>
		<TD BGCOLOR="#FFFFCC" COLSPAN = 3 align=right><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE"><a HREF="v2AdminEdit.asp?id=0&PBox=<%=menuid%>&Add=Y">Add SIDE item&nbsp;</FONT></TD>
	</TR>
	
	<TR>
		<TD BGCOLOR="#FFFFCC" align=center><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE"><B>edit</B></FONT></TD>
		<TD BGCOLOR="#FFFFCC" align=center><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE">category</FONT></TD>
		<TD BGCOLOR="#FFFFCC" align=center><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE">subsort</FONT></TD>
		<TD BGCOLOR="#FFFFCC" align=center><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE">category</FONT></TD>
		<TD BGCOLOR="#FFFFCC" align=center><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE">show</FONT></TD>
		</TR>
	<%DO WHILE NOT rs.EOF
	
		SQLcs = "select * from catsort where MainMenuid = " & menuid & " order by SideMenuSort;"
		set rscs = Server.CreateObject("ADODB.Recordset")
		rscs.Open sqlcs, Conn, adOpenKeyset, adLockOptimistic
		%>
		<form method=POST action="v2AdminEntry.asp?" name=x<%=rs("id")%>>
	
   		<TR>
			<TD><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE"><a HREF="v2AdminEdit.asp?id=<%=rs("id")%>&PBox=<%=menuid%>"><%=trim(RS("Title"))%>&nbsp;</FONT></TD>
			<TD><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="BLUE">			
			<SELECT	name='CatsortSelectBox' size='1' title='CatsortSelectTitle' onChange="document.x<%=rs("id")%>.action='v2AdminEntry.asp?PBox=<%=menuid%>&CBox=<%=rs("id")%>';document.x<%=rs("id")%>.submit();">
				<%DO WHILE NOT rscs.EOF
					if rs("catid") = rscs("id") then
						ThisIsItcs =rscs("id")
						strSelected = "SELECTED"
					else
						strSelected = ""
					end if%>
	
					<OPTION  <%=strSelected%> NAME=CatsortSelect value=<%=rscs("id")%>><%=RScs("SideMenuText")%></option>
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
			<% '<TD><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" COLOR="RED"><a HREF="v2AdminSideDelete.asp?id=<%=rs("id")%> <% '&PBox=<%=menuid%> <% ' "> DELETE</FONT></TD> %>
		</TR>
		</FORM>
		<%rs.movenext
	loop%> 
	
</TABLE>

<table align=center cellpadding=6>
	<TR>
		<TD>
			<a href="v2AdminMenuMaintenance.asp?PBox=<%=ThisIsIt%>">Change Menu</a>
		</td>
	</tr>
</table>

		
</p>

					
<%'*************************************
					%>
				</TD></TR>
			</TABLE>
		</td>
	</tr>		
</table>
</FONT>
</body>
</html>
