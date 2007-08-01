<!--#INCLUDE FILE="adovbs.inc"-->
<%
    ' IF Not IsEmpty(Request.form("id")) THEN 
		menuid = Request.Form("menuid")
	'else
	'	id = 0
	'end if

    IF Session("New") THEN 
		width = Request.Form("width")
		height = Request.Form("height")
		offset = Request.Form("offset")
	END IF

	Menuitem = Request.Form("menuitem")
	' On Error Resume Next
' build an sql statement to retrieve all fields

' Determine if we are going to add or edit
Addit = Request.Form("Addit")


'IF (Addit <> "Y") THEN

   sql = "SELECT * FROM Menuitems WHERE menuID = " & menuid & " ;"


'ELSE
'   sql = "SELECT * FROM Main where id = 0;"
'END IF

' response.write(sql)
' create connection objects
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.connectionstring = Session("connstring")
Conn.Open 	
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.Open sql, Conn, adOpenKeyset, adLockOptimistic
'IF Addit = "D" THEN
'  RS.Delete
'  RS.Close
'   set RS = nothing
'   set Conn = nothing

'   Response.Redirect ("AdminEntry.asp?mid=" & menuid & "")
'END IF

' clear errors
'Conn.Errors.Clear
'Err = 0

' Error check values
'IF Len(newProjectName) = 0 THEN newProjectName = Null

' If AEType = "Add" add a record to file, otherwise just replace the variables
'	IF (Addit = "Y") THEN 
'		RS.AddNew
'		RS("MenuID") = MenuID
'	END IF
		' Give record set fields their values

        RS("menuitem")  = menuitem
       IF session("New") Then
	           RS("height")  = height
	           RS("width")  = width
	           RS("offset")  = offset
	   
	   End IF
'IF (AEType = "Add") THEN
'		RS("StartDate")     = Date()
'END IF
        RS.Update ' Write values to record set

'If Err = 0 Then
'    Session("msg") = "Record added successfully!"
'    ' Response.Redirect "fullrec.asp?rec_no=" & rec_no
'Else
'  '  Session("msg") = "Unable to store the record.  Error: " & CStr(Err)
'  '  Response.Redirect "addnew.asp"
'End If

' close connection
Conn.Close

' free memory
set RS = nothing
set Conn = nothing

' Create a starter Subproject/Task for a new project
'IF AEType = "Add" THEN

'          sql = "SELECT * FROM Subprojects;"

        ' create connection objects
'         Set Conn = Server.CreateObject("ADODB.Connection")
 '        Set RS = Server.CreateObject("ADODB.RecordSet")

        ' open the connection
'          Conn.Open Session("ConnName")
 '         RS.Open sql, Conn, adOpenKeyset, adLockOptimistic

'        RS.AddNew

'        RS("Description") = CSTR("None")
'        RS("Priority")    = 0
'        RS("Start")       = CDATE(DATE)
'        RS("Due")         = CDATE(DATE)
'        RS("ProjectKey")  = ProjectKey
'        RS.UpDate
'        RS.Close

'END IF

Response.Redirect ("AdminEntry.asp?mid=" & menuid & "")
%>