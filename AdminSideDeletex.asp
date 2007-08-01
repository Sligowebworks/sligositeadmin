<!--#INCLUDE FILE="adovbs.inc"-->
 <%
	IF Not IsEmpty(Request.QueryString("mid")) THEN 
		menuid = request.querystring("mid")	
	else
		menuid = 0
	end if
	IF Not IsEmpty(Request.QueryString("id")) THEN 
		id = request.querystring("id")	
	else
		id = 0
	end if
	sql = "SELECT * FROM Main WHERE ID = " & id & " ;"
  	Set Conn = Server.CreateObject("ADODB.Connection")
	Set RS = Server.CreateObject("ADODB.RecordSet")
	' open the connection
	Conn.Open "learnersupport"
	RS.Open sql, Conn, adOpenKeyset, adLockOptimistic
	RS.Delete
  	RS.Close
   	set RS = nothing
  	set Conn = nothing

   	Response.Redirect ("AdminEntry.asp?mid=" & menuid & "")

%>