<!--#include file="ADOVBS.INC"-->
<%
zThisPage = Request.ServerVariables("PATH_INFO")
%>

 <!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!-- copyright 2001 Sligo Computer Services, Inc. All rights reserved -->
<html>
<head>
<title>Summer Camps 2001</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<meta name="keywords" content="computer services, web design, software development, programming, consultants, visual basic,
	school improvement, assessment, Maryland, D.C. Metropolitan area, VA, statistical analysis, data analysis, visual data presentation, networks, desktop publishing, layout,sligo">
	<meta name="description" content="Sligo Computer Services Web Site">	
</HEAD>
<%
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.connectionstring = Session("connstring")
	Conn.Open 	
	%>
<!--#INCLUDE FILE="functionsv2.inc"-->
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
<DIV ALIGN="CENTER">
<!--<BODY BACKGROUND="Images/bg4.gif">-->
<body BGCOLOR="white">

	
<% 	logosource = "http://Collabitat.com/SligoSite/" + trim(session("path")) + "/Images/Logo.gif"
	dotsource =  "http://Collabitat.com/SligoSite/" + trim(session("path")) + "/Images/dot.gif"
	homesource = "http://Collabitat.com/SligoSite/" + trim(session("path")) + "/Index.asp"
%>
<% ' The table for the image and top menu starts here %>
<table WIDTH="600" BORDER="0">
	<tr>
		<td HEIGHT="33" BGCOLOR="white" ALIGN="middle" COLSPAN="2"><img src="<%= Logosource %>" BORDER="0"> </td>
	</tr>	
	<tr>
		<td ALIGN="RIGHT" VALIGN="bottom" >
		<P id="topredlink">
			<%PaintTopMenu() %>
		</P>		
		</td>		 
	</tr>
	<tr>
		<td COLSPAN="2" BGCOLOR="#0066ff"><IMG height=2 src="Images/Dot.gif"></td>
	</tr>
</table>
<% ' now for the bottom of the page. First the big table %>
<table WIDTH="600" cellpadding="5" BORDER="0">
	<tr>
		<td WIDTH="145" BGCOLOR="#0066ff" HEIGHT="450" ALIGN="left" VALIGN="top" >
		<TABLE WIDTH="140" ALIGN="left"  BORDER="0">

		<% ' This is the side menu table%>
			<tr>
				<td>
					<P id="colhead">		
						<%PaintSideTitle()%>
						<br>
					
					</P>
				</td>
			</tr>
						<%PaintSideMenu()%>	

<!--<img WIDTH="140" HEIGHT="135" src="Images/CityMap.gif"><BR>-->

		</TABLE>

		</td>
		<% ' Now comes the area for content %>
		<td width=430 VALIGN="top" BACKGROUND="Images/bg1.gif" >
			<TABLE ALIGN="JUSTIFY" WIDTH="430" CELLPADDING="5"  BORDER="0">
				<TR><TD>						
					<%'PaintContent()%>
					
					<!--#include file="v2AdminEntry.asp"-->
					
				</TD></TR>
			</TABLE>
		</td>	
	</tr>  		
</table>
<DIV></DIV></FONT>
</body>
</html>
