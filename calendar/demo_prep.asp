<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 10/28/1999

' set appropriate permissions and redirect to calendar

if Request.Form("level") <> "" then
	Session("status") = CInt(Request.Form("level"))
	response.redirect "webCal4_month.asp"
end if
%>

<html>
<body bgcolor="#c0c0c0">
<font face="Tahoma, Arial, Helvetica" size=2>
<form action="demo_prep.asp" method="post">

<input type="submit" value="Set Access Level">
<select name="level">
<option>0
<option>1
<option>2
<option>3
<option>4
<option>5
</select>

</form>
<br>
Zero (0) is guest access while values greater than zero will provide varying permission levels for staff (not yet implemented).
</body>
</html>
