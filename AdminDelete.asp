<%


IF Not IsEmpty(Request.QueryString("id")) THEN 
		id = request.querystring("id")	
	'	%>id:<%=id%><%
	Set Conn = Server.CreateObject("ADODB.Connection")
	' get conn string from session
	' USE FOR DEVT:
	' Conn.connectionstring  = "Provider=SQLOLEDB.1;Password=bytes4us;Persist Security Info=True;User ID=sa;Initial Catalog=CDHP;Data Source=Inferno2"
	' USE FOR STARFISH:
	' Conn.connectionstring = "Provider=SQLOLEDB.1;Password=bytes4us;Persist Security Info=True;User ID=sa;Initial Catalog=SligoSite;Data Source=STARFISH"
	Conn.connectionstring = Session("connstring")
	Conn.Open
	'----------------------------------------------------------------------------------------------------------------------------------------------------------------	
	'Deleting file Add by Bagus
	'sqlquery="SELECT MenuName,PageName FROM Main WHERE ID=" & id & " ;"
	'set rsquery= Server.CreateObject("ADODB.Recordset")
	'rsquery.Open sqlquery,Conn,adOpenKeyset,adLockOptimistic
	'rsquery.MoveFirst
	'dim objFSO
	'Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
    'vfile = "D:\Websites\SligoSite\CDHP\" & RS("menuname") & "\"  & RS("PageName") & ".asp"
	'objFSO.DeleteFile vfile
	set rsquery=nothing
	'----------------------------------------------------------------------------------------------------------------------------------------------------------------		
	sql = "DELETE FROM Main WHERE ID = " & id & " ;"
	set rsdel = Server.CreateObject("ADODB.Recordset")
	set rsm = Conn.execute(sql)
	'rsdel.Open sql, Conn, adOpenKeyset, adLockOptimistic
	set rsdel=nothing
	Conn.Close
end if
Response.Redirect ("AdminEntry.asp?id=" & id & "")
%>
