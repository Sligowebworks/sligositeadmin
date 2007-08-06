<!--#INCLUDE FILE="adovbs.inc"-->
<% 
IF request.querystring("PW") <> 1 then
	Session("ErrMessage") = ""
	end if
	%>
<HTML><HEAD><TITLE>SligoSiteAdminEntry 51805</TITLE></HEAD>
<!-- copyright 1997-2005 Sligo Computer Services, Inc. All rights reserved -->
<BODY TEXT="#6c6c6c" LINK="#484848" VLINK=#5C5C5C" onLoad="document.PasswordForm.Password.focus()">
 &nbsp;&nbsp;&nbsp; 
<div align="center">
<TABLE width="580" align="center" border=0>
	<TR>
		<TD height="50" ALIGN="MIDDLE">
			<FONT FACE="verdana" size="+1" COLOR="#006699">
			<B>SligoSite Administrator Entry Page</B></FONT>
		</TD> 
	</TR>
	<TR>
		<TD align="center" COLSPAN="2"><div align="center">
	        <% IF Session("ErrMessage") <> "" THEN %>
                <FONT COLOR="RED" FACE="Arial, Helvetica" ><%=Session("ErrMessage")%></FONT>
        	<%ELSE%>
                <FONT FACE="Arial, Helvetica" >
                This site is for Administrators of SligoSite - <br> people with passwords.
                </FONT>
        	<%END IF%>
	

<!-- #INCLUDE FILE="db.connection.asp" -->
<%
sql="SELECT SiteName FROM Admin;"
Set rsOrgs = Server.CreateObject("ADODB.Recordset")
rsOrgs.Open sql, Conn, adOpenKeyset, adLockOptimistic	
%>		
		</TD>
	</TR>
</TABLE>
<TABLE width=580 align=center>
	<TR>
		<TD ALIGN="RIGHT">Organization:</TD>
		<TD>
        	<FORM Name="PasswordForm" Action="security.asp" METHOD="POST">
        		<SELECT Name="WhichOrganization">
				<%DO WHILE NOT RSOrgs.EOF%>
					<OPTION> <%=rsOrgs("SiteName")%>
					<%rsOrgs.movenext%>
				<%Loop%>	
				</SELECT>
				<% rsOrgs.Close %>
				<BR>
	            <FONT FACE="Arial, Helvetica" >
		</TD>
	</TR>
	<TR>
		<TD ALIGN="RIGHT">Password:</TD>
		<TD>
				<INPUT TYPE="PASSWORD" NAME="Password" SIZE=20></FONT>
		</TD>
	</TR>
    <TR>
		<TD ALIGN="CENTER" COLSPAN="2" height=170>          
                <INPUT TYPE="SUBMIT" VALUE=" -- CONTINUE -- ">
		 	</FORM>
		</TD>
	</TR><BR><BR>	
	<TR>
		<TD align="center" COLSPAN="2">
			<FONT Size=-1><B>A project of</B></FONT>  <FONT SIZE=-1 Color="#006699"><B> Sligo Computer Services, Inc.</B></Font>
		</TD>
	</TR>
</TABLE>
	
</DIV></DIV>
</BODY>
