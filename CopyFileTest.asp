<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
</head>
<body>
 <%
    dim objFSO
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	vfile = "D:\Websites\SligoSite\CDHP\Test\template2.asp"
    response.Write(vfile)
    objFSO.CopyFile "D:\Websites\SligoSite\CDHP\Template\" &  "template.asp",vfile
%> 

</body>
</html>