<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 10/28/1999

' the field names in the two tables are currently
' different so we need two sets of variables

dim strEventQuery, strProgQuery, intFacility, strType
dim strCombine, intEventFacility, strEventType

if Request.Form("facility") <> "" then

' if the form was selected then update session variables

	intFacility = Request.Form("facility")
	strType = Request.Form("type")
	Session("intEventFacility") = intFacility
	Session("strEventType") = strType

elseif Session("intEventFacility") <> "" then

' or if session variables already exist then assign to local
' values
	
	intFacility = Session("intEventFacility")
	strType = Session("strEventType")

end if

if intFacility <> "" then

' if we end up with local values, one way or another,
' generate new query

	if intFacility <> 0 then
		strEventQuery = " AND (Location=" & intFacility
		strCombine = " AND "
	else
		strEventQuery = " "
		strCombine = " AND ("
	end if

	Select Case strType
		Case "all"
			if strEventQuery <> " " then
				strEventQuery = strEventQuery & ") "
			end if
			strProgQuery = " "
		Case "rec"
			strEventQuery = strEventQuery & strCombine & "Recreation=" & True & ") "
			strProgQuery = " AND (RecProgram=" & True & ") "
		Case "ca"
			strEventQuery = strEventQuery & strCombine & "CulturalArts=" & True & ") "
			strProgQuery = " AND (CAProgram=" & True & ") "
		Case "after"
			strEventQuery = strEventQuery & strCombine & "AfterSchool=" & True & ") "
			strProgQuery = " AND (AfterHoursProgram=" & True & ") "
	End Select
else

' if there was no form selection and no session values then
' use blanks

	strEventQuery = " "
	strProgQuery = " "
end if

' these will become part of the query we use to pull data
' from the events and programs tables
%>