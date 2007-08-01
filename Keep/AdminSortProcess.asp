<!--#INCLUDE FILE="adovbs.inc"-->

<%
Set Conn = Server.CreateObject("ADODB.Connection")
	' get conn string from session
	' session("connstring") = "Provider=SQLOLEDB.1;Password=zenmind;Persist Security Info=True;User ID=sa;Initial Catalog=SligoSite_LearnerSupport;Data Source=COLLABITAT"
	' Conn.connectionstring = "Provider=SQLOLEDB.1;Password=zenmind;Persist Security Info=True;User ID=sa;Initial Catalog=SligoSite_LearnerSupport;Data Source=COLLABITAT"
	Conn.connectionstring = Session("connstring")
	Conn.Open 	
	
%>
<%  IF Not IsEmpty(Request.QueryString("mid")) THEN 
		mmid = request.querystring("mid")	
	else
		mmid = 0
	end if
	  IF Not IsEmpty(Request.QueryString("sid")) THEN 
		sid = request.querystring("sid")	
	else
		sid = 0
	end if
	' Conn.cursorlocation = aduseclient
	
	SQLset = "select * from Main where Menuid = " & mmid & " order by CatSort, sort;"
	set rs = Server.CreateObject("ADODB.Recordset")
	' rs.Open sqlset, Conn, adOpenKeyset, adLockOptimistic
	rs.Open sqlset, Conn, adOpenKeyset, adLockOptimistic
	
	%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Data Entry Form</title>
</head>

	
<body>
hi
</body>
</html>

