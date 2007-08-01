<!-- ASP Script -->
<!-- Style is here -->
	<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Data Entry Form</title>
</head>

<body >

<%
response.Write(request.Form("title"))
response.Write(request.QueryString("PageName"))
%>
</body>
</html>

