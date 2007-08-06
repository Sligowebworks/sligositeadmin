<%
Set Conn = Server.CreateObject("ADODB.Connection")

Conn.connectionstring = "Provider=SQLOLEDB.1;Password=bytes4us;Persist Security Info=True;User ID=sa;Initial Catalog=SligoSite;Data Source=local"

Conn.Open 
%>		
