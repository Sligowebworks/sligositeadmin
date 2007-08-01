<!--#INCLUDE FILE="adovbs.inc"-->
<%
    ' IF Not IsEmpty(Request.form("id")) THEN 
		id = Request.Form("id")	
		menuid = Request.Form("mid")
	'else
	'	id = 0
	'end if
       	title = Request.Form("title")
        tophead = Request.Form("tophead")
        bottomhead = Request.Form("bottomhead")
        graphic = Request.Form("graphic")
        category = Request.Form("category")
        catsort  = Request.Form("catsort")
        sort = Request.Form("sort")
        noshow = Request.Form("noshow")
        description = Request.Form("description")
' On Error Resume Next
' build an sql statement to retrieve all fields

' Determine if we are going to add or edit
Addit = Request.Form("Addit")

IF (Addit <> "Y") THEN
   sql = "SELECT * FROM Main WHERE ID = " & id & " ;"
ELSE
   sql = "SELECT * FROM Main where id = 0;"
END IF

'response.write(sql)
' create connection objects
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.connectionstring = Session("connstring")
	Conn.Open 	
	set rs = Server.CreateObject("ADODB.Recordset")
	RS.Open sql, Conn, adOpenKeyset, adLockOptimistic
	
IF Addit = "D" THEN
  RS.Delete
  RS.Close
   set RS = nothing
   set Conn = nothing

   Response.Redirect ("AdminEntry.asp?mid=" & menuid & "")
END IF

' clear errors
'Conn.Errors.Clear
'Err = 0

' Error check values
'IF Len(newProjectName) = 0 THEN newProjectName = Null

' If AEType = "Add" add a record to file, otherwise just replace the variables
	IF (Addit = "Y") THEN 
		RS.AddNew
		RS("MenuID") = MenuID
	END IF
		' Give record set fields their values

        RS("title")    = title
        RS("tophead") = tophead
        RS("bottomhead") = bottomhead
        RS("graphic") = graphic
        RS("category") = category
        RS("catsort")  = catsort
        if sort = "" then 
			sort = 0
		end if 
		RS("sort") = sort
		
        if noshow = "on" then
			RS("noshow") = -1
		else
			rs("noshow") = 0
		end if
        RS("description") = description




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