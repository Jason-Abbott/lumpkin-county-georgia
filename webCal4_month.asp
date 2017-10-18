<html>
<head>
<!--#include file="webCal4_themes.inc"-->
<!--#include file="webCal4_rollovers.inc"-->
<SCRIPT TYPE="text/javascript" LANGUAGE="javascript" SRC="webCal4_buttons.js"></SCRIPT>
<!--#include file="webCal4_data.inc"-->
<!--#include file="webCal4_define-month.inc"-->
<!--#include file="webCal4_option-parse.inc"-->
<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 10/28/1999

' ---------------------------------------------------------
' build an array of event data for selected month
' ---------------------------------------------------------
' calendar items come from two tables with different
' field names, so separate queries:
' FIRST query the events table
	
strQuery = "SELECT * FROM tblEvents WHERE (Date " _
	& strCommon & ")" & strEventQuery & "ORDER BY Start"
	
' put all matching events in an array indexed by day number

Set rsEvents = Server.CreateObject("ADODB.RecordSet")

'  cursor: adOpenStatic = 3
' locking: adLockReadOnly = 1
' command: adCmdText = &H0001

rsEvents.Open strQuery, strDSN, 3, 1, &H0001

do while not rsEvents.EOF
	intIndex = Day(rsEvents("Date"))
	
	strDescription = Replace(TimeValue(rsEvents("Start")), ":00 ", " ") _
		& " to " & Replace(TimeValue(rsEvents("End")), ":00 ", " ")

	arEvents(intIndex) = arEvents(intIndex) _
		& "<img src=""./images/arrow_right_black" _
		& ".gif"" width=4 height=7> " & VbCrLf _
		& "<a href=""webCal4_detail.asp?event_id=" & rsEvents("EventID") _
		& "&date=" & rsEvents("Date") & "&view=month"" " _
		& showStatus(strDescription) & ">" _
		& rsEvents("Title") & "</a><br>" & VbCrLf
	rsEvents.MoveNext
loop

rsEvents.Close
set rsEvents = nothing

' SECOND, query the programs table

strQuery = "SELECT * FROM tblPrograms WHERE (StartDate " _
	& strCommon & " OR EndDate " & strCommon & ")" _
	& strProgQuery
	
' put all matching programs in an array indexed by day number

Set rsPrograms = Server.CreateObject("ADODB.RecordSet")
rsPrograms.Open strQuery, strDSN, 3, 1, &H0001
do while not rsPrograms.EOF

' determine if the program ends or begins in this month

	if Month(rsPrograms("StartDate")) = m then
		intIndex = Day(rsPrograms("StartDate"))
		strName = "Start of " & rsPrograms("ProgramName")
	else
		intIndex = Day(rsPrograms("EndDate"))
		strName = "End of " & rsPrograms("ProgramName")
	end if

' different programs have their own detail pages
' select appropriate link
	
	if rsPrograms("RecProgram") = True then
		strLink = "RecProgramDetail"
	elseif rsPrograms("CAProgram") = True then
		strLink = "CAProgramDetail"
	elseif rsPrograms("AfterHoursProgram") = True then
		strLink = "AfProgramDetail"
	end if
	
	arEvents(intIndex) = arEvents(intIndex) _
		& "<img src=""./images/arrow_right_green" _
		& ".gif"" width=4 height=7> " & VbCrLf _
		& "<a href=""../" & strLink & ".asp?id=" & rsPrograms("ProgramID") _
		& "&view=month"" " _
		& showStatus(rsPrograms("Participants")) & ">" _
		& strName & "</a><br>" & VbCrLf

	rsPrograms.MoveNext
loop

rsPrograms.Close
set rsPrograms = nothing

' create the text to display at the top of the calendar

strTitle = "<font face=""Verdana, Arial, Helvetica"" color=""#" _
	& arColor(6) & """ size=4><b>DEMO</b> (access: " & Session("status") & ")</font>"
%>

</head>
<body bgcolor="#<%=arColor(1)%>" link="#<%=arColor(7)%>" vlink="#<%=arColor(7)%>" alink="#<%=arColor(6)%>">

<table width="100%" border=0 cellspacing=0 cellpadding=1>
<!--#include file="webCal4_buttons.inc"-->
<tr>
	<td bgcolor="#<%=arColor(6)%>" align="center" colspan=2>
<!--#include file="webCal4_layout-month.inc"-->
	</td>
</table>

<center>
<!--#include file="webCal4_option-form.inc"-->
</center>

<font face="Verdana, Arial, Helvetica" size=1>
<a href="http://boise.uidaho.edu/jason/webCal.html" target="_top">
webCal 4.0</a>
</font>

</body>
</html>