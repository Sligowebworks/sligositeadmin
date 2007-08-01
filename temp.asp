<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Untitled</title>
</head>

<body>
<%
	IF Not IsEmpty(Request.Form("PBox")) THEN 
		%>PBOX = <%=Request.Form("PBox")%><br><%
		PBox=Request.Form("PBox")
	else
		%>PBOX = Nada<br><%
	end if
	
	
	IF Not IsEmpty(Request.Form("CBox")) THEN 
		%>CBOX = <%=Request.Form("CBox")%><br><%
		CBox=Request.Form("CBox")
	else
		%>CBOX = Nada<br><%
	end if
	
	IF Not IsEmpty(Request.Form("ZBox")) THEN 
		%>ZBOX = <%=Request.Form("ZBox")%><br><%
		ZBox=Request.Form("ZBox")
	else
		%>ZBOX = Nada<br><%
	end if
	
	%>
	
<table>
<form method=POST name="FormTest" action="temp.asp">
	<TR>
		<TD>	
			<SELECT	name='PBox' size='1' title='PBoxTitle' onChange="document.FormTest.action='temp.asp';document.FormTest.submit();">
			<%IF PBox = 1 then%>
					<OPTION SELECTED NAME=Select1 value=1>PBOX Item 1</option>
					<OPTION 		 NAME=Select1 value=2>PBOX Item 2</option>
			<%ELSE%>
					<OPTION 		 NAME=Select1 value=1>PBOX Item 1</option>
					<OPTION SELECTED NAME=Select1 value=2>PBOX Item 2</option>
			<%END IF%>
			</select>
		</TD>
	<!--/FORM-->
<!--form method=POST action="v2AdminMenuMaintenance.asp?" name="CategorySelect"-->
		<TD>	
			<SELECT	name='CBox' size='1' title='CBoxTitle' onChange="document.FormTest.action='temp.asp';document.FormTest.submit();">
			<%IF CBox = 1 then%>
					<OPTION SELECTED NAME=Select2 value=1>CBOX Item 1</option>
					<OPTION 		 NAME=Select2 value=2>CBOX Item 2</option>
			<%ELSE%>
					<OPTION 		 NAME=Select2 value=1>CBOX Item 1</option>
					<OPTION SELECTED NAME=Select2 value=2>CBOX Item 2</option>
			<%END IF%>
			</select>
		</TD>
<%if CBox = 1 then %>
	</tr>
	</FORM>
<%else%>
		<TD>	
			<SELECT	name='ZBox' size='1' title='zBoxTitle' onChange="document.FormTest.action='temp.asp';document.FormTest.submit();">
			<%IF ZBox = 1 then%>
					<OPTION SELECTED NAME=Select2 value=1>ZBOX Item 1</option>
					<OPTION 		 NAME=Select2 value=2>ZBOX Item 2</option>
			<%ELSE%>
					<OPTION 		 NAME=Select2 value=1>ZBOX Item 1</option>
					<OPTION SELECTED NAME=Select2 value=2>ZBOX Item 2</option>
			<%END IF%>
			</select>
		</TD>
	</tr>
	</FORM>
<%end if%>
</TABLE>

<script language="javascript">
function ChangeCurrPage()
	{
		// zVal = "?PBox=" ;
		// zVal = zVal + document.FormPageSelect.PBox.options[document.FormPageSelect.PBox.selectedIndex].value ; 
		// document.FormPageSelect.action = "v2AdminMenuMaintenance.asp" + zVal ;
		document.FormPageSelect.submit();
	}
	
</script>

</body>
</html>
