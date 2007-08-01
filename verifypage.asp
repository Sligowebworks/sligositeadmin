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
 <table WIDTH="720" height=80% BORDER="0" align="center">
	<tr>
		<td BGCOLOR="white" align=center valign=middle >
		<%
		    response.Write ("The Page Name You Entered Already Exist. Please use different page name.<p>")
		    response.Write ("<a href=AdminEntry.asp>Back to Admin Page</a>") 
		
		%>
		</td>
	</tr>
</TABLE>
</body>
</html>