<!--#INCLUDE FILE="adovbs.inc"-->
<%if request.Querystring("ORG") = "Sligo" and request.Querystring("PW") <> "" then
		WhichOrg = "SligoSite"
		Session("WhichOrganization") = "Sligo Computer Services, Inc."
		TestPassword = UCASE(Request.Querystring("PW")) ' <- User password
else             
		WhichOrg = LTRIM(Request.Form("WhichOrganization"))
		Session("WhichOrganization") = Request.Form("WhichOrganization")
		TestPassword = UCASE(Request.FORM("Password")) ' <- User password
end if

Set Conn = Server.CreateObject("ADODB.Connection")
Conn.connectionstring = "Provider=SQLOLEDB.1;Password=bytes4us;Persist Security Info=True;User ID=sa;Initial Catalog=SligoSite;Data Source=COLLABITAT"
Conn.Open 	

sql = "SELECT * FROM Admin WHERE SiteName = '" & TRIM(WhichOrg) & "' ;"

Set rsOrgs = Server.CreateObject("ADODB.Recordset")
rsOrgs.Open sql, Conn, adOpenKeyset, adLockOptimistic
session("connstring") = rsOrgs("ConnectString")
session("path") = rsOrgs("path")
session("sitename") = rsOrgs("sitename")

If TrIM(TestPassword) = "" then
	rsOrgs.Close
    Session("ErrMessage") = "You must enter a password to access this site."
    RESPONSE.REDIRECT("Index.asp?PW=1")
ELSEIf TrIM(TestPassword) = UCASE(TRIM(rsOrgs("Password"))) then
	rsOrgs.Close
    RESPONSE.REDIRECT("Adminindexv2.asp")
ELSE
	rsOrgs.Close
    Session("ErrMessage") = "Sorry incorrect password.  Please try again"
    RESPONSE.REDIRECT("Indexv2.asp?PW=1")
END IF 

%>