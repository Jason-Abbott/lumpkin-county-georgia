<table bgcolor="#<%=arColor(12)%>">
<form action="webCal4_<%=strView%>.asp?date=<%=strThisDay%>" method="post">
<tr>
	<td align="center"><font face="Tahoma, Arial, Helvetica" size=2>
<input type="submit" value="View">

<select name="facility">
<option value=0>All Facilities
<%
dim rsFac

strQuery = "SELECT * FROM tblFacilities ORDER BY Name"
Set rsFac = Server.CreateObject("ADODB.RecordSet")
rsFac.Open strQuery, strDSN, 3, 1, &H0001
do while not rsFac.EOF
	response.write "<option value=""" & rsFac("FacilityID") & """"
	if CInt(rsFac("FacilityID")) = CInt(intFacility) then
		response.write " selected"
	end if
	response.write ">" & rsFac("Name") & VbCrLf
	rsFac.MoveNext
loop
rsFac.Close
Set rsFac = nothing
%>
</select>, 

<select name="type">
<option value="all">All
<option value="rec"<%if strType = "rec" then%> selected<%end if%>>Recreation
<option value="ca"<%if strType = "ca" then%> selected<%end if%>>Cultural Arts
<option value="after"<%if strType = "after" then%> selected<%end if%>>After Hours
</select>
programs

	</font>
	</td>
</form>
</table>