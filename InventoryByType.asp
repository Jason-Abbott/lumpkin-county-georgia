<%@ Language=VBScript%>
<%
programtype=request("searchtype")
response.write programtype
dim SQL
x=0

Set Conn = Server.CreateObject("ADODB.Connection")

Conn.Open "lumpkin"

set rs = Server.CreateObject("ADODB.Recordset")

SQL = "Select * from tblEquipmentType WHERE InventoryType="&request("searchtype")

rs.open SQL, conn, 1, 2


'Response.Write "<DIV style='position:absolute;left:155;top:250;width:771;z-index:12'>"

'Response.Write"<table border='1' width='100%'>"

'  Response.Write"<tr>"
'    Response.Write"<td width='20%'><strong>Equipment Type</strong></td>"
'    Response.Write"<td width='20%'><strong>Description</strong></td>"
'	Response.Write"<td width='10%'><strong>Checked Out</strong></td>"
'	Response.Write"<td width='20%'><strong>Program</strong></td>"
'	Response.Write"<td width='10%'><strong>Facility</strong></td>"
'	Response.Write"<td width='10%'><strong>Staff</strong></td>"
'	Response.Write"<td width='10%'><strong>Non Staff</strong></td>"
'  Response.Write"</tr>"

'Do While Not rs.EOF

'	Response.Write "<tr>"
'	Response.Write "<td>"
'	Response.Write rs("Type")
'	Response.Write "</td><td>"
'	Response.Write rs("Description")
'	Response.Write "</td><td>"
'		Set Conn = Server.CreateObject("ADODB.Connection")
'		Conn.Open "lumpkin"
'		set rs2 = Server.CreateObject("ADODB.Recordset")
'		SQL = "Select * from tblEquipment where Type="&request("searchtype")
'		rs2.open SQL, conn, 1, 2
'		Response.Write rs2("CheckedOut")
'		Response.Write "</td><td>"
'		Response.Write rs2("ProgramID")
'		Response.Write "</td><td>"
'		Response.Write rs2("FacilityID")
'		Response.Write "</td><td>"		
'		Response.Write rs2("StaffID")
'		Response.Write "</td><td>"
'		Response.Write rs2("NonStaffID")
'		Response.Write "</td>"
'		rs2.close
'		conn.close
'		set conn=nothing

'	rs.MoveNext
	
'Loop 

'	Response.Write "</tr></table>"
'	Response.Write "</DIV>"

rs.close
conn.close
set conn = nothing
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" CONTENT="Elemental Software Drumbeat 2000 3.01.369">

<TITLE>Page39</TITLE>

<SCRIPT SRC="media/DBRelPos.js" LANGUAGE="JavaScript"></SCRIPT>

<!-- CSS styles and positioning for IE 4 browsers. -->
<LINK REL="StyleSheet" TYPE="text/css" HREF="Lumpkin_IE4.css">
<STYLE ID="IE4_ABSOLUTE_POSITIONING_STYLE_SHEET">
#EndOfPage { position:absolute;left:0;top:10;width:0 }
</STYLE>

</HEAD>
<BODY TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0   BACKGROUND="media/green.gif" BGColor="#ffffff">

<A NAME="StartOfPage"></A>


<DIV ID=Section0 CLASS=section STYLE="position:relative; top:0; height:0">
</DIV><SCRIPT>__calcSectionHeight(Section0, 0)</SCRIPT>

<DIV ID="EndOfPage"><A NAME="EndOfPage"></A></DIV>

</BODY></HTML>
