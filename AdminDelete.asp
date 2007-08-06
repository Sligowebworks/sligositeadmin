<%

IF Not IsEmpty(Request.QueryString("id")) THEN 
	id = request.querystring("id")	

	Set Conn = Server.CreateObject("ADODB.Connection")
	' get conn string from session
	Conn.connectionstring = Session("connstring")
	Conn.Open

	sql = "DELETE FROM Main WHERE ID = " & id & " ;"
	set rsdel = Server.CreateObject("ADODB.Recordset")
	set rsm = Conn.execute(sql)

End If
If NOT IsEmpty(Request.Form("PageName")) Then
	dim objFSO
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

	Path = objFSO.BuildPath(objFSO.BuildPath("d:\websites\sligosite\", Session("path") ), Request.Form("MenuName"))
	FileName = Request.Form("PageName") & ".asp"
	If objFSO.FileExists(objFSO.BuildPath(Path, FileName)) Then
		call objFSO.DeleteFile(objFSO.BuildPath(Path, FileName ), True)
	End If

	set objFSO = nothing

End If

Response.Redirect ("AdminEntry.asp?id=" & id & "")
%>
