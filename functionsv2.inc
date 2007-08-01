<%
if request.querystring("Type") <> "" then
	SStype = "I"
else
	SStype = "I"
end if

if request.querystring("ID") <> "" then
	SSID = 0
else
	SSID = 0
end if

if request.querystring("SecID") <> "" then
	secID = request.querystring("SecID")
else
	SecID = 0
end if

SQLsec = "select * from Menuitems order by sort;"
set rsm = Server.CreateObject("ADODB.Recordset")
rsm.Open sqlsec, Conn, adOpenStatic, adLockReadOnly 
'response.write sqlsec

SQLsec2 = "select * from Menuitems where Menuid = " & SecID & " ;"
set rsm2 = Server.CreateObject("ADODB.Recordset")
rsm2.Open sqlsec2, Conn, adOpenStatic, adLockReadOnly 

SectionName = rsm2("MenuItem")

function PaintTopMenu()
	
	'write the section menu
	Do while not rsm.eof
	%><%=rsm("MenuID")%><%
		SecName = rsm("MenuItem")
		response.write "<A HREF=AdminIndexv2.asp?Type=S&SecID=" & rsm("MenuID") & " >" & SecName & "</A>&nbsp;&nbsp;"
		rsm.movenext
	loop
end function
	
function PaintSideTitle()
	response.write SectionName
end function

function PaintSideMenu()
	' SET THE PA SELECT
	'SQLside = "select * from Main where Menuid = " & SecID & " order by Sort, Catsort ;"
	SQLside = "select Main.ID, Catsort.ID as catID, Main.Title, Catsort.SideMenuText as Category, Main.Sort, Catsort.SideMenuSort from Main " _
				& " inner Join Catsort on main.catid = Catsort.id where Menuid = " & SecID & " order by Catsort.SideMenuSort, Main.Sort;"
	set rsside = Server.CreateObject("ADODB.Recordset")
	rsside.Open SQLside, Conn, adOpenStatic, adLockReadOnly 
	'Response.Write "<br>" & SQLside & "<br>"
	rsside.movefirst
	CatID = "1"
	DO WHILE NOT rsside.EOF 
		'aMenuCode(PAGEID,zPA) = rsside("ID")
		'strQueryString = ""
		'For iLoop = 0 to kount -1
    	'  strQueryString = strQueryString & "XX=" & aMenuCode(PAGEID,iLoop)
    	'  strQueryString = strQueryString & "&"
		'next
		'Cut off the last character
		'strQueryString = Left(strQueryString, Len(strQueryString) - 1)
		' PRINT NEW CATEGORIES
		
		IF CatID <> rsside("Category") THEN
			response.write "<tr><td ALIGN=LEFT VALIGN=TOP>" _
							& "<a href=""AdminIndexv2.asp?Type=C&SecID=" & SecID & "&ID=" & rsside("CatID") & """><P id=cathead>" & rsside("Category") _
							& "</P></a></td></tr>"
			CatID = rsside("Category")
		END IF
		response.write "<tr><td ALIGN=LEFT VALIGN=TOP>"
		' REDDEN CURRENT CHOICE "#990033"
		IF CINT(SSID) = CINT(rsside("ID")) THEN
			response.write "<P id=redlink>" 
		ELSE
			response.write "<P id=sidelink>"
		END IF
		response.write "<A HREF=""AdminIndexv2.asp?" & "Type=I&SecID=" & SecID & "&ID=" & rsside("ID") & """>" & rsside("Title") & "</A>"
		response.write "</P></td></tr>"
		rsside.movenext
	LOOP
end function

function PaintContent()
	' Displays stuff on right
	SQL = "SELECT * FROM Main where MenuID = " & CINT(zPA) & " and noshow <> -1 order by CatSort, sort;"
	set rs = conn.execute(sql)	
	RS.movefirst
	DO WHILE NOT RS.EOF
		IF CINT(zID) = CINT(RS("ID")) THEN
			'check for Desc - print it if it's there
			IF RS("TopHead") <> "" Then%>
				<P id="tophead"><%=RS("TopHead")%></P>
			<%END IF
			IF RS("BottomHead") <> "" Then%>
				<P id="bottomhead"><%=RS("BottomHead")%></P>
			<%END IF
		
			IF RS("Graphic") <> "" Then%>					
				<img ALIGN="LEFT" VSPACE="5" HSPACE="20" src="<%=RS("Graphic")%>">
			<%END IF%>
		
			<P align="justify" id="basicstyle"><%=RS("Description")%></P>
		<%END IF
		rs.movenext
	LOOP
	set rs = nothing
	set conn = nothing		
end function
%>
<script type="text/javascript" language="JavaScript">
function opennew(zurl) {
window.open(zurl, "newwindow", "status,Height=400,WIDTH=750,alwaysRaised,resizable,scrollbars,screenx=100,screeny=50")
}
</script>