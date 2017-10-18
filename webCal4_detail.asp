<html>
<head>
<script language="javascript"><!--
//preload images and text for faster operation

if (document.images) {
// back to calendar icon
	var iconMonth = new Image();
	iconMonth.src = "images/icon_calprev_grey.gif";
	var iconMonthOn = new Image();
	iconMonthOn.src = "images/icon_calprev.gif"
}

//-->
</script>
<!--#include file="webCal4_rollovers.inc"-->
<!--#include file="webCal4_themes.inc"-->
<!--#include file="webCal4_data.inc"-->

<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 10/28/1999

dim rs, url, strDescription, strRecur, strTitle, intID
dim strDate, strStart, strEnd, strView, strType, tfStaff

' these are the variables used by the included webCal4_showrecur:

dim arMonths(12), strYear, x, objYears

strView = Request.QueryString("view")
strCal = Request.QueryString("type")
	
' pull the event information and dates from db

Session("programnum") = Request.QueryString("event_id")
	
strQuery = "SELECT * FROM tblEvents E INNER JOIN tblFacilities F " _
	& "ON E.Location = F.FacilityID " _
	& "WHERE (EventID)=" & Request.QueryString("event_id")

Set rsEvent = Server.CreateObject("ADODB.RecordSet")

'  cursor: adOpenStatic = 3
' locking: adLockReadOnly = 1
' command: adCmdText = &H0001

rsEvent.Open strQuery, strDSN, 3, 1, &H0001
	
	if rsEvent("Desc") <> "" then
		strDetail = "<dt><b>Description</b>" & VbCrLf & "<dd>" _
			& Replace(rsEvent("Desc"), VbCrLf, "<br>") & "<p>"
	end if
	
	if rsEvent("Name") <> "" then
		strDetail = strDetail & "<dt><b>Location</b>" & VbCrLf & "<dd>" _
			& Replace(rsEvent("Name"), VbCrLf, "<br>") & "<p>"
	end if

	if rsEvent("Notes") <> "" then
		strDetail = strDetail & "<dt><b>Notes</b>" & VbCrLf & "<dd>" _
			& Replace(rsEvent("Notes"), VbCrLf, "<br>") & "<p>"
	end if

	strTitle = rsEvent("Title")
	intID = rsEvent("EventID")
	strDate = Request.QueryString("date")
	strStart = TimeValue(rsEvent("Start"))
	strEnd = TimeValue(rsEvent("End"))
rsEvent.Close
Set rsEvent = nothing

%>

</head>
<body bgcolor="#<%=arColor(1)%>" link="#<%=arColor(7)%>" vlink="#<%=arColor(7)%>" alink="#<%=arColor(6)%>">
<br>
<center>
<table border=0 cellspacing=0 cellpadding=3>
<tr>
	<td rowspan=2 align="center" valign="top">
		<font face="Tahoma, Arial, Helvatica" color="#<%=arColor(4)%>">
		<b><font size=2><%=WeekdayName(WeekDay(strDate))%></font><br>
		<font size=7><%=Day(strDate)%></font><br>
		<font size=5><%=MonthName(Month(strDate),1)%></font></b><br>
		<font size=4><%=Year(strDate)%></font>
		</font>
	</td>

	<td valign="top">
	<table cellspacing=0 cellpadding=3 border=0 width="100%">
	<tr>
		<td bgcolor="#<%=arColor(4)%>">
			<font face="Verdana, Arial, Helvetica" size=4><b>
			<a href="webCal4_<%=strView%>.asp?date=<%=eventDate%>"
			<%=switchIcon("Month","","Return to Calendar")%>><img name="Month" src="./images/icon_calprev_grey.gif" width=15 height=16 alt="" border=0></a>

			<%=strTitle%></b></font>&nbsp;
			
		</td>
<tr>
	<td><font face="Tahoma, Arial, Helvetica" size=2>
	<dl><%=strDetail%></dl></font>
	</td>
</table>

<%
' if this is a recurring event then get list of all
' dates on which it occurs

' recurrence isn't presently supported so make it always false

if 1 = 4 then
' if strRecur > 0 then
	dim intCount, arDates()

	strQuery = "SELECT * FROM cal_dates" _
		& " WHERE event_id=" & intID _
		& " ORDER BY event_date"

	Set rsDates = Server.CreateObject("ADODB.RecordSet")

'  cursor: adOpenForwardOnly = 0
' locking: adLockReadOnly = 1
' command: adCmdText = &H0001

	rsDates.Open strQuery, strDSN, 0, 1, &H0001

' generate array of event dates

	intCount = 0
	do while not rsDates.EOF
		ReDim preserve arDates(intCount)
		arDates(intCount) = rsDates("event_date")
		intCount = intCount + 1
		rsDates.MoveNext
	loop
	rsDates.Close
	Set rsDates = nothing
	
' if the event recurs then display the dates on
' which it occurs, invoking the include file
' that does the special formatting
%>
	<table bgcolor="#<%=arColor(2)%>" width="100%"><form><tr><td>
	<!--#include file="webCal4_showrecur.inc"-->
	</td></form></table>
<%
end if
response.write "</td>" & VbCrLf

'-----------------------------------
' display time range if one was entered for this event
'-----------------------------------

if strStart <> "" then
	dim intHrStart, intHrEnd, intSpan, strHrCurrent, strHrColor, strTxtColor

' the Hour function formats to military time
	
	intHrStart = Hour(strStart)
	intHrEnd = Hour(strEnd)
	
' we need to handle those events that END on midnight, which actually
' is the start of the next day

	if intHrEnd = 0 then
		intHrEnd = 24
	end if

	response.write "<td rowspan=2 valign=""top"">" _
		& "<table cellspacing=1 cellpadding=0 border=0>"

' calculate the hours spanned by the event

	intSpan = (strHrEnd - strHrStart) + 1

	for h = 0 to 24
		if h = intHrStart then
			strHrCurrent = "<b>" & Replace(strStart, ":00 ", " ") & "</b>"
		elseif h = intHrEnd then
			strHrCurrent = "<b>" & Replace(strEnd, ":00 ", " ") & "</b>"
		else

' otherwise insert the array value with regular clock notation
' appended, changing 12PM to noon for the temporally challenged
			
			if h = 0 OR h = 24 then
				strHrCurrent = "<b>midnight</b>"
			elseif h < 12 then
				strHrCurrent = h & ":00 AM"
			elseif h = 12 then
				strHrCurrent = "<b>noon</b>"
			else
				strHrCurrent = h - 12 & ":00 PM"
			end if
		end if

' make the hours covered by the event a different color
		
		if h >= intHrStart AND h <= intHrEnd then
			strHrColor = arColor(9)
			strTxtColor = arColor(6)
		else
			strHrColor = arColor(2)
			strTxtColor = arColor(6)
		end if

		response.write "<tr><td bgcolor=""#" & strHrColor _
			& """ align=""right"" nowrap><font face=""Tahoma, Arial, Helvetica""" _
			& "size=1 color=""#" & strTxtColor & """>" _
			& strHrCurrent & "</font></td>"
	next
	response.write "</td></table>"
end if

' from here display the management buttons
%>

<!-- buttons -->

<% if Session("status") > 0 then %>

<tr>
	<td valign="bottom">
	
	<!-- framing table -->
	<table bgcolor="#<%=arColor(5)%>" width="100%" cellspacing=0 cellpadding=2 border=0><tr><td>
	<!-- end framing table -->
	
	<table cellspacing=0 cellpadding=2 border=0 width="100%">
	<tr>
		<form action="../ManageEvents.asp" method="post">
		<td bgcolor="#<%=arColor(12)%>">
		<input type="submit" name="edit" value="Manage">
		</td>
		</form>
		
		<form action="../CancelEvent.asp" method="post">
		<td bgcolor="#<%=arColor(12)%>" align="right">
		<input type="submit" name="delete" value="Cancel">
		</td>
		</form>
	</table>

	<!-- framing table -->
	</td></table>
	<!-- end framing table -->
	
<% end if %>

	</td>
</table>

</center>
</body>
</html>